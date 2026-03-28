Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrape()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, logPath As String, vbsCode As String
    Dim sName As String
    
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    logPath = "C:\temp\trace.log"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath
    If fso.FileExists(logPath) Then fso.DeleteFile logPath

    ' --- BUILD THE VBSCRIPT ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, r, txt, detail, buf, rows, cols, clist, i, sName, regEx, lastRow, exitLoop" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "regEx.Global = True : regEx.Pattern = ""\s{2,}""" & vbCrLf
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' Session Connection Logic
    vbsCode = vbsCode & "Set clist = CreateObject(""PCOMM.autECLConnList"") : clist.Refresh" & vbCrLf
    vbsCode = vbsCode & "sName = """" : For i = 1 To clist.Count" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, clist(i).Name, ""BTS Pooled"", 1) > 0 Then sName = clist(i).Name : Exit For" & vbCrLf
    vbsCode = vbsCode & "Next : If sName = """" Then sName = clist(1).Name" & vbCrLf
    
    vbsCode = vbsCode & "Set objSess = CreateObject(""PCOMM.autECLSession"") : objSess.SetConnectionByName(sName)" & vbCrLf
    vbsCode = vbsCode & "Set objPS = objSess.autECLPS : Set objOIA = objSess.autECLOIA" & vbCrLf
    vbsCode = vbsCode & "rows = objPS.NumRows : cols = objPS.NumCols" & vbCrLf
    
    ' MAIN PAGE LOOP
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    vbsCode = vbsCode & "  objPS.SetCursorPos 7, 80" & vbCrLf ' Start just above the data
    vbsCode = vbsCode & "  lastRow = 0" & vbCrLf
    
    ' TABBING LOOP (Scans current page underscores)
    vbsCode = vbsCode & "  Do" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[tab]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 150" & vbCrLf ' Sufficient time for cursor to move
    
    ' Exit tab loop if cursor leaves the 8-22 range or jumps back to top
    vbsCode = vbsCode & "    If objPS.CursorRow < 8 Or objPS.CursorRow > 22 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    If objPS.CursorRow <= lastRow Then Exit Do" & vbCrLf
    
    vbsCode = vbsCode & "    lastRow = objPS.CursorRow : r = objPS.CursorRow" & vbCrLf
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf
    
    ' --- DRILL DOWN DANCE ---
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf2]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 600" & vbCrLf 
    ' Grab Detail (Change 5, 10, 20 to your specific field coordinates)
    vbsCode = vbsCode & "    detail = Trim(objPS.GetText(5, 10, 20))" & vbCrLf 
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf3]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 600" & vbCrLf 
    
    ' Save to CSV
    vbsCode = vbsCode & "    txt = regEx.Replace(txt, ""|"")" & vbCrLf
    vbsCode = vbsCode & "    objFile.WriteLine txt & ""|"" & detail" & vbCrLf
    vbsCode = vbsCode & "  Loop" & vbCrLf ' End Tab Loop
    
    ' --- TERMINATION CHECK ---
    vbsCode = vbsCode & "  buf = UCase(objPS.GetText(1, 1, rows * cols))" & vbCrLf
    ' If "Invalid" or "Last Page" appears anywhere, we are done
    vbsCode = vbsCode & "  If InStr(1, buf, ""INVALID"") > 0 Or InStr(1, buf, ""LAST PAGE"") > 0 Then Exit Do" & vbCrLf
    
    ' Move to next page
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep 1000" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close"

    ' --- EXECUTE ---
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    On Error Resume Next
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    On Error GoTo 0

    ' --- IMPORT ---
    If fso.FileExists(csvPath) Then ImportData csvPath
    If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath
End Sub

Sub ImportData(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited: .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .Refresh BackgroundQuery:=False
    End With
    ws.Range("A1:J1").Value = Array("STAT", "ACCOUNT", "BRKR", "O/S ACCT", "RR", "DATE", "AGE", "RFT ID", "PLAN", "DETAIL_INFO")
    ws.Columns("A").Replace "_ ", ""
    On Error Resume Next: ws.Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete: On Error GoTo 0
    ws.UsedRange.Columns.AutoFit
End Sub
