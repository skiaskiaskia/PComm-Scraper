Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrape()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, vbsCode As String
    
    ' File Paths
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    ' Initial Cleanup
    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath

    ' --- BUILD THE VBSCRIPT BRIDGE ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, r, txt, detail, rows, cols, clist, i, sName, regEx, lastRow, fRow, fCol, buf, rStart, matches" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' PComm Session Connection Logic
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
    
    ' Look for the first row starting with "_ XXXX"
    vbsCode = vbsCode & "  buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "  regEx.Pattern = ""_ [A-Z]{3,4}"" : regEx.Global = False : regEx.IgnoreCase = True" & vbCrLf
    vbsCode = vbsCode & "  Set matches = regEx.Execute(buf)" & vbCrLf
    vbsCode = vbsCode & "  If matches.Count > 0 Then" & vbCrLf
    vbsCode = vbsCode & "     i = matches(0).FirstIndex + 1" & vbCrLf
    vbsCode = vbsCode & "     rStart = ((i - 1) \ cols) + 1" & vbCrLf
    vbsCode = vbsCode & "     objPS.SetCursorPos rStart - 1, cols" & vbCrLf ' Start cursor at end of row above data
    vbsCode = vbsCode & "  Else" & vbCrLf
    vbsCode = vbsCode & "     rStart = 8 : objPS.SetCursorPos 7, cols" & vbCrLf ' Fallback if pattern fails
    vbsCode = vbsCode & "  End If" & vbCrLf
    
    vbsCode = vbsCode & "  lastRow = 0" & vbCrLf
    
    ' --- TABBING LOOP (PER PAGE) ---
    vbsCode = vbsCode & "  Do" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[tab]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 200" & vbCrLf 
    
    ' Exit tab loop if cursor leaves the data area or jumps back to top
    vbsCode = vbsCode & "    If objPS.CursorPosRow < rStart Or objPS.CursorPosRow > 22 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    If objPS.CursorPosRow <= lastRow Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    lastRow = objPS.CursorPosRow : r = objPS.CursorPosRow" & vbCrLf
    
    ' Capture Main Row Text
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf
    
    ' --- DRILL DOWN TO F2 ---
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf2]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 800" & vbCrLf 
    
    ' Search for SPEC INS: and grab 4 rows inline
    vbsCode = vbsCode & "    detail = ""None""" & vbCrLf
    vbsCode = vbsCode & "    buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "    i = InStr(1, buf, ""SPEC INS:"", 1)" & vbCrLf
    vbsCode = vbsCode & "    If i > 0 Then" & vbCrLf
    vbsCode = vbsCode & "       fRow = ((i - 1) \ cols) + 1" & vbCrLf
    vbsCode = vbsCode & "       fCol = ((i - 1) Mod cols) + 1 + 9" & vbCrLf ' Offset by length of "SPEC INS:"
    vbsCode = vbsCode & "       detail = """"" & vbCrLf
    vbsCode = vbsCode & "       For i = 0 To 3" & vbCrLf ' Grab row found + 3 rows below
    vbsCode = vbsCode & "          If (fRow + i) <= rows Then detail = detail & Trim(objPS.GetText(fRow + i, fCol, cols - fCol + 1)) & "" """ & vbCrLf
    vbsCode = vbsCode & "       Next" & vbCrLf
    vbsCode = vbsCode & "       detail = Trim(detail)" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    
    ' Return to main list
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf11]""" & vbCrLf 
    vbsCode = vbsCode & "    WScript.Sleep 800" & vbCrLf 
    
    ' Format row with Pipes for CSV
    vbsCode = vbsCode & "    regEx.Pattern = ""\s{2,}|\|"" : regEx.Global = True" & vbCrLf
    vbsCode = vbsCode & "    txt = regEx.Replace(txt, ""|"")" & vbCrLf
    vbsCode = vbsCode & "    detail = Replace(detail, ""|"", "" "")" & vbCrLf ' Prevent internal pipes from breaking columns
    vbsCode = vbsCode & "    objFile.WriteLine txt & ""|"" & detail" & vbCrLf
    vbsCode = vbsCode & "  Loop" & vbCrLf 
    
    ' --- PAGING LOGIC & CIRCUIT BREAKER ---
    vbsCode = vbsCode & "  buf = UCase(objPS.GetText(1, 1, rows * cols))" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, buf, ""INVALID"") > 0 Or InStr(1, buf, ""LAST PAGE"") > 0 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep 1200" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close"

    ' --- EXECUTE VBSCRIPT ---
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    On Error Resume Next
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    On Error GoTo 0

    ' --- IMPORT TO EXCEL ---
    If fso.FileExists(csvPath) Then ImportData csvPath
    ' Cleanup bridge file
    If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath
    
    MsgBox "Scrape Complete!", vbInformation
End Sub

Sub ImportData(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.ClearContents
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited: .TextFileOtherDelimiter = "|"
        ' Setting 12 columns as Text (2) to prevent dropping leading zeros on account numbers
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .Refresh BackgroundQuery:=False
    End With
    
    ' Headers
    ws.Range("A1:K1").Value = Array("STAT", "ACCOUNT", "BRKR", "O/S ACCT", "RR", "DATE", "AGE", "RFT ID", "PLAN", "SPEC_INSTRUCTIONS")
    
    ' Final cleanup of STAT underscores
    ws.Columns("A").Replace "_ ", ""
    
    ' Remove any orphan blank rows
    On Error Resume Next: ws.Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete: On Error GoTo 0
    ws.UsedRange.Columns.AutoFit
End Sub
