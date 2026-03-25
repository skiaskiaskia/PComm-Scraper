Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrape()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, logPath As String, vbsCode As String
    Dim sName As String, rowCount As Long
    
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    logPath = "C:\temp\trace.log"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    ' Ensure directory exists and clean old files
    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath
    If fso.FileExists(logPath) Then fso.DeleteFile logPath

    ' --- BUILD THE VBSCRIPT PIECE BY PIECE ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, objLog, r, txt, rows, cols, clist, i, sName, regEx, lastScrn, currScrn, wait" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "regEx.Global = True : regEx.Pattern = ""\s{2,}""" & vbCrLf
    vbsCode = vbsCode & "Sub UpdateLog(msg) : Set objLog = objFSO.CreateTextFile(""" & logPath & """, True) : objLog.WriteLine(msg) : objLog.Close : End Sub" & vbCrLf
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' Session Discovery
    vbsCode = vbsCode & "Set clist = CreateObject(""PCOMM.autECLConnList"") : clist.Refresh" & vbCrLf
    vbsCode = vbsCode & "sName = """" : For i = 1 To clist.Count" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, clist(i).Name, ""BTS Pooled"", 1) > 0 Then sName = clist(i).Name : Exit For" & vbCrLf
    vbsCode = vbsCode & "Next : If sName = """" Then sName = clist(1).Name" & vbCrLf
    
    ' Connection Setup
    vbsCode = vbsCode & "Set objSess = CreateObject(""PCOMM.autECLSession"") : objSess.SetConnectionByName(sName)" & vbCrLf
    vbsCode = vbsCode & "Set objPS = objSess.autECLPS : Set objOIA = objSess.autECLOIA" & vbCrLf
    vbsCode = vbsCode & "rows = objPS.NumRows : cols = objPS.NumCols" & vbCrLf
    
    ' MAIN DO LOOP
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    vbsCode = vbsCode & "  lastScrn = objPS.GetText(1, 1, 100)" & vbCrLf ' Snapshot for sync
    
    ' Scrape Rows 8 to 22
    vbsCode = vbsCode & "  For r = 8 To 22" & vbCrLf 
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf 
    vbsCode = vbsCode & "    txt = regEx.Replace(txt, ""|"")" & vbCrLf 
    vbsCode = vbsCode & "    If Len(txt) > 5 Then objFile.WriteLine txt" & vbCrLf
    vbsCode = vbsCode & "  Next" & vbCrLf
    
    ' Exit condition check
    vbsCode = vbsCode & "  If InStr(1, UCase(objPS.GetText(1, 1, rows * cols)), ""LAST PAGE"") > 0 Then Exit Do" & vbCrLf
    
    ' Advance Page (PA1)
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf 
    
    ' Smart Sync Wait
    vbsCode = vbsCode & "  For wait = 1 To 20" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 100" & vbCrLf
    vbsCode = vbsCode & "    currScrn = objPS.GetText(1, 1, 100)" & vbCrLf
    vbsCode = vbsCode & "    If currScrn <> lastScrn Then Exit For" & vbCrLf
    vbsCode = vbsCode & "  Next" & vbCrLf
    
    ' LOOP END
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close"

    ' --- RUN THE BRIDGE ---
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    On Error Resume Next
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    On Error GoTo 0

    ' --- IMPORT INTO EXCEL ---
    If fso.FileExists(csvPath) Then
        ImportData csvPath
        rowCount = ThisWorkbook.Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row - 1
        MsgBox "Scrape Complete! " & rowCount & " rows imported.", vbInformation
    Else
        MsgBox "No data found. Ensure PComm is on the correct screen.", vbExclamation
    End If
    
    ' Cleanup temp bridge
    If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath
End Sub

Sub ImportData(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    
    ' Connect to CSV
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited
        .TextFileOtherDelimiter = "|"
        ' Format first 10 columns as Text to keep leading zeros
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .Refresh BackgroundQuery:=False
    End With
    
    ' Cleanup connections
    Dim qt As QueryTable: For Each qt In ws.QueryTables: qt.Delete: Next
    
    ' Format Sheet
    ws.Range("A1:I1").Value = Array("STAT", "ACCOUNT", "BRKR", "O/S ACCT", "RR", "DATE", "AGE", "RFT ID", "PLAN")
    ws.Range("A1:I1").Font.Bold = True
    
    ' Clean leading underscores often found in STAT columns
    ws.Columns("A").Replace "_ ", ""
    
    ' 1-line cleanup: delete empty rows and autofit
    On Error Resume Next
    ws.Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    ws.UsedRange.Columns.AutoFit
    On Error GoTo 0
End Sub
