Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrape()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, vbsCode As String
    
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath

    ' --- BUILD THE VBSCRIPT ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, r, txt, detail, rows, cols, clist, i, sName, regEx, lastRow, fRow, fCol" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "regEx.Global = True : regEx.Pattern = ""\s{2,}|\|""" & vbCrLf 
    
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' Session Connection
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
    vbsCode = vbsCode & "  objPS.SetCursorPos 7, 80" & vbCrLf 
    vbsCode = vbsCode & "  lastRow = 0" & vbCrLf
    
    ' TABBING LOOP
    vbsCode = vbsCode & "  Do" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[tab]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 200" & vbCrLf 
    
    vbsCode = vbsCode & "    If objPS.CursorPosRow < 8 Or objPS.CursorPosRow > 22 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    If objPS.CursorPosRow <= lastRow Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    lastRow = objPS.CursorPosRow : r = objPS.CursorPosRow" & vbCrLf
    
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf
    
    ' --- DRILL DOWN ---
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf2]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 800" & vbCrLf 
    
    vbsCode = vbsCode & "    detail = ""None""" & vbCrLf
    vbsCode = vbsCode & "    fRow = 1 : fCol = 1" & vbCrLf ' Initial search position
    vbsCode = vbsCode & "    If objPS.SearchText(""SPEC INS:"", 1, fRow, fCol) Then" & vbCrLf
    vbsCode = vbsCode & "       fCol = fCol + 9" & vbCrLf 
    vbsCode = vbsCode & "       detail = """"" & vbCrLf
    vbsCode = vbsCode & "       For i = 0 To 3" & vbCrLf
    vbsCode = vbsCode & "          detail = detail & Trim(objPS.GetText(fRow + i, fCol, cols - fCol + 1)) & "" """ & vbCrLf
    vbsCode = vbsCode & "       Next" & vbCrLf
    vbsCode = vbsCode & "       detail = Trim(detail)" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf11]""" & vbCrLf 
    vbsCode = vbsCode & "    WScript.Sleep 800" & vbCrLf 
    
    ' Cleanup & Write
    vbsCode = vbsCode & "    txt = regEx.Replace(txt, ""|"")" & vbCrLf
    vbsCode = vbsCode & "    detail = Replace(detail, ""|"", "" "")" & vbCrLf
    vbsCode = vbsCode & "    objFile.WriteLine txt & ""|"" & detail" & vbCrLf
    vbsCode = vbsCode & "  Loop" & vbCrLf 
    
    ' PAGING & TERMINATION
    vbsCode = vbsCode & "  If InStr(1, UCase(objPS.GetText(1, 1, rows * cols)), ""INVALID"") > 0 Or InStr(1, UCase(objPS.GetText(1, 1, rows * cols)), ""LAST PAGE"") > 0 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep 1000" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close"

    ' --- EXECUTION ---
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True

    ' --- IMPORT ---
    If fso.FileExists(csvPath) Then ImportData csvPath
End Sub

Sub ImportData(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited: .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .Refresh BackgroundQuery:=False
    End With
    ws.Range("A1:J1").Value = Array("STAT", "ACCOUNT", "BRKR", "O/S ACCT", "RR", "DATE", "AGE", "RFT ID", "PLAN", "SPEC_INSTRUCTIONS")
    ws.Columns("A").Replace "_ ", ""
    On Error Resume Next: ws.Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete: On Error GoTo 0
    ws.UsedRange.Columns.AutoFit
End Sub
