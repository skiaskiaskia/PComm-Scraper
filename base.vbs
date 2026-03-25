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
    
    ' 1. SETUP PATHS & CLEANUP
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    logPath = "C:\temp\trace.log"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath
    If fso.FileExists(logPath) Then fso.DeleteFile logPath

    ' 2. BUILD THE VBSCRIPT (WITH SMART SYNC & REGEX)
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, objLog, r, txt, rows, cols, clist, i, sName, regEx, lastScrn, currScrn, wait" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "regEx.Global = True : regEx.Pattern = ""\s{2,}""" & vbCrLf
    
    vbsCode = vbsCode & "Sub UpdateLog(msg) : Set objLog = objFSO.CreateTextFile(""" & logPath & """, True) : objLog.WriteLine(msg) : objLog.Close : End Sub" & vbCrLf
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' Find Session
    vbsCode = vbsCode & "Set clist = CreateObject(""PCOMM.autECLConnList"") : clist.Refresh" & vbCrLf
    vbsCode = vbsCode & "sName = """" : For i = 1 To clist.Count" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, clist(i).Name, ""BTS Pooled"", 1) > 0 Then sName = clist(i).Name : Exit For" & vbCrLf
    vbsCode = vbsCode & "Next : If sName = """" Then sName = clist(1).Name" & vbCrLf
    vbsCode = vbsCode & "UpdateLog ""SESSION:"" & sName" & vbCrLf
    
    ' Connect
    vbsCode = vbsCode & "Set objSess = CreateObject(""PCOMM.autECLSession"") : objSess.SetConnectionByName(sName)" & vbCrLf
    vbsCode = vbsCode & "Set objPS = objSess.autECLPS : Set objOIA = objSess.autECLOIA" & vbCrLf
    vbsCode = vbsCode & "rows = objPS.NumRows : cols = objPS.NumCols" & vbCrLf
    
    ' SCRAPE LOOP
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    vbsCode = vbsCode & "  lastScrn = objPS.GetText(1, 1, 100)" & vbCrLf
    
    vbsCode = vbsCode & "  For r = 8 To 22" & vbCrLf 
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf 
    vbsCode = vbsCode & "    txt = regEx.Replace(txt, ""|"")" & vbCrLf 
    vbsCode = vbsCode & "    If Len(txt) > 5 Then objFile.WriteLine txt" & vbCrLf
    vbsCode = vbsCode & "  Next" & vbCrLf
    
    ' Exit Check
    vbsCode = vbsCode & "  If InStr(1, UCase(objPS.GetText(1, 1, rows * cols)), ""LAST PAGE"") > 0 Then Exit Do" & vbCrLf
    
    ' Smart Advance: Click and Wait for Change
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf 
    vbsCode = vbsCode & "  For wait = 1 To 20" & vbCrLf ' Wait up to 2 seconds
    vbsCode = vbsCode & "    WScript.Sleep 100" & vbCrLf
    vbsCode = vbsCode & "    currScrn = objPS.GetText(1, 1, 100)" & vbCrLf
    vbsCode = vbsCode & "    If currScrn <> lastScrn Then Exit For" & vbCrLf
    vbsCode = vbsCode & "  Next" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close : UpdateLog ""DONE""" & vbCrLf

    ' 3. RUN THE BRIDGE
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    On Error Resume Next
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    
    ' 4. GET SESSION NAME FOR MSG
    If fso.FileExists(logPath) Then
        Dim ts As Object, ln As String
        Set ts = fso.OpenTextFile(logPath, 1)
        Do Until ts.AtEndOfStream: ln = ts.ReadLine: If Left(ln, 8) = "SESSION:" Then sName = Mid(ln, 9): Loop
        ts.Close
    End If

    ' 5. IMPORT & CLEAN
    If fso.FileExists(csvPath) Then
        ImportData csvPath
        rowCount = ThisWorkbook.Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row - 1
        MsgBox "Cooked! Scraped " & rowCount & " rows from Session: " & sName, vbInformation
    Else
        MsgBox "Scrape failed to generate data. Check PComm window.", vbCritical
    End If
    
    If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath
End Sub

Sub ImportData(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited: .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .Refresh BackgroundQuery:=False
    End With
    
    ' Cleanup Connections
    Dim qt As QueryTable: For Each qt In ws.QueryTables: qt.Delete: Next
    
    ' Headers & Underscore Strip
    ws.Range("A1:I1").Value = Array("STAT", "ACCOUNT", "BRKR", "O/S ACCT", "RR", "DATE", "AGE", "RFT ID", "PLAN")
    ws.Range("A1:I1").Font.Bold = True
    ws.Columns("A").Replace "_ ", "" ' Clean up the leading underscores
    
    ' The 1-line row wiper & format
    On Error Resume Next
    ws.Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    ws.UsedRange.Columns.AutoFit
    On Error GoTo 0
End Sub
