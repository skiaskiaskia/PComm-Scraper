Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrapeRaw()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, vbsCode As String
    
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath

    ' --- BUILD THE VBSCRIPT BRIDGE ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, r, txt, detail, rows, cols, clist, i, sName, regEx, lastRow, fRow, fCol, buf, rStart, matches, lastScreen" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' Session Connection
    vbsCode = vbsCode & "Set clist = CreateObject(""PCOMM.autECLConnList"") : clist.Refresh" & vbCrLf
    vbsCode = vbsCode & "sName = """" : For i = 1 To clist.Count" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, clist(i).Name, ""BTS Pooled"", 1) > 0 Then sName = clist(i).Name : Exit For" & vbCrLf
    vbsCode = vbsCode & "Next : If sName = """" Then sName = clist(1).Name" & vbCrLf
    
    vbsCode = vbsCode & "Set objSess = CreateObject(""PCOMM.autECLSession"") : objSess.SetConnectionByName(sName)" & vbCrLf
    vbsCode = vbsCode & "Set objPS = objSess.autECLPS : Set objOIA = objSess.autECLOIA" & vbCrLf
    vbsCode = vbsCode & "rows = objPS.NumRows : cols = objPS.NumCols" & vbCrLf
    
    ' --- MAIN PAGE LOOP ---
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    vbsCode = vbsCode & "  buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "  lastScreen = buf" & vbCrLf 
    
    ' Dynamic Start Search (_ XXXX)
    vbsCode = vbsCode & "  regEx.Pattern = ""_ [A-Z]{3,4}"" : regEx.Global = False : regEx.IgnoreCase = True" & vbCrLf
    vbsCode = vbsCode & "  Set matches = regEx.Execute(buf)" & vbCrLf
    vbsCode = vbsCode & "  If matches.Count > 0 Then" & vbCrLf
    vbsCode = vbsCode & "     i = matches(0).FirstIndex + 1" & vbCrLf
    vbsCode = vbsCode & "     rStart = ((i - 1) \ cols) + 1" & vbCrLf
    vbsCode = vbsCode & "     objPS.SetCursorPos rStart - 1, cols" & vbCrLf
    vbsCode = vbsCode & "  Else" & vbCrLf
    vbsCode = vbsCode & "     rStart = 8 : objPS.SetCursorPos 7, cols" & vbCrLf
    vbsCode = vbsCode & "  End If" & vbCrLf
    
    vbsCode = vbsCode & "  lastRow = 0" & vbCrLf
    
    ' --- TABBING LOOP ---
    vbsCode = vbsCode & "  Do" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[tab]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 200" & vbCrLf 
    vbsCode = vbsCode & "    If objPS.CursorPosRow < rStart Or objPS.CursorPosRow > 22 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    If objPS.CursorPosRow <= lastRow Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    lastRow = objPS.CursorPosRow : r = objPS.CursorPosRow" & vbCrLf
    
    ' RAW SCRAPE: No Trim, No RegEx cleanup
    vbsCode = vbsCode & "    txt = objPS.GetText(r, 1, cols)" & vbCrLf
    
    ' --- DRILL DOWN ---
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf2]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 850" & vbCrLf 
    
    vbsCode = vbsCode & "    detail = """"" & vbCrLf
    vbsCode = vbsCode & "    buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "    i = InStr(1, buf, ""SPEC INS:"", 1)" & vbCrLf
    vbsCode = vbsCode & "    If i > 0 Then" & vbCrLf
    vbsCode = vbsCode & "       fRow = ((i - 1) \ cols) + 1 + 1" & vbCrLf 
    vbsCode = vbsCode & "       fCol = ((i - 1) Mod cols) + 1 + 9" & vbCrLf 
    vbsCode = vbsCode & "       For i = 0 To 1" & vbCrLf 
    vbsCode = vbsCode & "          If (fRow + i) <= rows Then" & vbCrLf
    vbsCode = vbsCode & "             detail = detail & objPS.GetText(fRow + i, fCol, cols - fCol + 1) & "" """ & vbCrLf
    vbsCode = vbsCode & "          End If" & vbCrLf
    vbsCode = vbsCode & "       Next" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf11]""" & vbCrLf 
    vbsCode = vbsCode & "    WScript.Sleep 850" & vbCrLf 
    
    ' NO DATA CLEANUP: Just escape existing pipes so they don't break the CSV
    vbsCode = vbsCode & "    txt = Replace(txt, ""|"", "" "")" & vbCrLf
    vbsCode = vbsCode & "    detail = Replace(detail, ""|"", "" "")" & vbCrLf
    vbsCode = vbsCode & "    objFile.WriteLine txt & ""|"" & detail" & vbCrLf
    vbsCode = vbsCode & "  Loop" & vbCrLf 
    
    ' --- PAGING & TERMINATION ---
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep 1200" & vbCrLf
    vbsCode = vbsCode & "  buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, UCase(buf), ""INVALID"") > 0 Or buf = lastScreen Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close"

    ' --- EXECUTION ---
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    On Error Resume Next
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    On Error GoTo 0

    ' --- EXCEL IMPORT ---
    If fso.FileExists(csvPath) Then ImportRawData csvPath
    If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath
    MsgBox "Raw Scrape Complete!", vbInformation
End Sub

Sub ImportRawData(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2, 2) ' Columns are strictly TEXT
        .Refresh BackgroundQuery:=False
    End With
    ws.Range("A1:B1").Value = Array("RAW_ROW_DATA", "RAW_SPEC_INS")
    ws.UsedRange.Columns.AutoFit
End Sub
