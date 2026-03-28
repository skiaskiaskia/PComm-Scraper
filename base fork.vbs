Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrapeClean()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, vbsCode As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    
    ' --- PULL SPEEDS FROM EXCEL CONTROL PANEL ---
    ' Change B1-B4 to match your preferred settings
    Dim sTab As Long: sTab = IIf(ws.Range("B1").Value > 0, ws.Range("B1").Value, 200)
    Dim sF2 As Long: sF2 = IIf(ws.Range("B2").Value > 0, ws.Range("B2").Value, 850)
    Dim sF11 As Long: sF11 = IIf(ws.Range("B3").Value > 0, ws.Range("B3").Value, 850)
    Dim sPage As Long: sPage = IIf(ws.Range("B4").Value > 0, ws.Range("B4").Value, 1200)
    
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")
    If Not fso.FolderExists("C:\temp") Then fso.CreateFolder ("C:\temp")
    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath

    ' --- BUILD THE VBSCRIPT ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, r, txt, detail, rows, cols, clist, i, sName, regEx, lastRow, fRow, fCol, buf, rStart, matches, lastScreen" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    
    ' [BENDER COMMENTED OUT]
    ' vbsCode = vbsCode & "Dim htaPath : htaPath = ""C:\temp\dancer.hta""" & vbCrLf
    ' vbsCode = vbsCode & "...HTA Window Creation Logic..." 
    
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    
    ' Session Connection
    vbsCode = vbsCode & "Set clist = CreateObject(""PCOMM.autECLConnList"") : clist.Refresh" & vbCrLf
    vbsCode = vbsCode & "sName = """" : For i = 1 To clist.Count" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, clist(i).Name, ""BTS Pooled"", 1) > 0 Then sName = clist(i).Name : Exit For" & vbCrLf
    vbsCode = vbsCode & "Next : If sName = """" Then sName = clist(1).Name" & vbCrLf
    vbsCode = vbsCode & "Set objSess = CreateObject(""PCOMM.autECLSession"") : objSess.SetConnectionByName(sName)" & vbCrLf
    vbsCode = vbsCode & "Set objPS = objSess.autECLPS : Set objOIA = objSess.autECLOIA" & vbCrLf
    vbsCode = vbsCode & "rows = objPS.NumRows : cols = objPS.NumCols" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    
    ' Main Loop
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    vbsCode = vbsCode & "  buf = objPS.GetText(1, 1, rows * cols) : lastScreen = buf" & vbCrLf
    vbsCode = vbsCode & "  regEx.Pattern = ""_ [A-Z]{3,4}"" : regEx.Global = False : regEx.IgnoreCase = True" & vbCrLf
    vbsCode = vbsCode & "  Set matches = regEx.Execute(buf)" & vbCrLf
    vbsCode = vbsCode & "  If matches.Count > 0 Then" & vbCrLf
    vbsCode = vbsCode & "     i = matches(0).FirstIndex + 1 : rStart = ((i - 1) \ cols) + 1" & vbCrLf
    vbsCode = vbsCode & "     objPS.SetCursorPos rStart - 1, cols" & vbCrLf
    vbsCode = vbsCode & "  Else : rStart = 8 : objPS.SetCursorPos 7, cols : End If" & vbCrLf
    vbsCode = vbsCode & "  lastRow = 0" & vbCrLf
    
    ' Tabbing / Scrape
    vbsCode = vbsCode & "  Do" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[tab]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep " & sTab & vbCrLf 
    vbsCode = vbsCode & "    If objPS.CursorPosRow < rStart Or objPS.CursorPosRow > 22 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    If objPS.CursorPosRow <= lastRow Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    lastRow = objPS.CursorPosRow : r = objPS.CursorPosRow" & vbCrLf
    vbsCode = vbsCode & "    txt = objPS.GetText(r, 1, cols)" & vbCrLf
    
    ' Drill Down
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf2]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep " & sF2 & vbCrLf 
    vbsCode = vbsCode & "    detail = """" : buf = objPS.GetText(1, 1, rows * cols) : i = InStr(1, buf, ""SPEC INS:"", 1)" & vbCrLf
    vbsCode = vbsCode & "    If i > 0 Then" & vbCrLf
    vbsCode = vbsCode & "       fRow = ((i - 1) \ cols) + 1 + 1 : fCol = ((i - 1) Mod cols) + 1 + 9" & vbCrLf 
    vbsCode = vbsCode & "       For i = 0 To 1" & vbCrLf 
    vbsCode = vbsCode & "          If (fRow + i) <= rows Then detail = detail & objPS.GetText(fRow + i, fCol, cols - fCol + 1) & "" "" " & vbCrLf
    vbsCode = vbsCode & "       Next" & vbCrLf
    vbsCode = vbsCode & "    End If" & vbCrLf
    
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf11]""" & vbCrLf 
    vbsCode = vbsCode & "    WScript.Sleep " & sF11 & vbCrLf 
    vbsCode = vbsCode & "    objFile.WriteLine Replace(txt, ""|"", "" "") & ""|"" & Replace(detail, ""|"", "" "")" & vbCrLf
    vbsCode = vbsCode & "  Loop" & vbCrLf 
    
    ' Paging
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep " & sPage & vbCrLf
    vbsCode = vbsCode & "  buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, UCase(buf), ""INVALID"") > 0 Or buf = lastScreen Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    vbsCode = vbsCode & "objFile.Close" & vbCrLf

    ' --- EXECUTE ---
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    
    ' --- IMPORT STARTING AT ROW 6 ---
    If fso.FileExists(csvPath) Then ImportRawDataToRow6 csvPath
    
    MsgBox "Scrape Complete! Data starts on Row 6.", vbInformation
End Sub

Sub ImportRawDataToRow6(path As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    
    ' Clear existing data from Row 5 down (preserves your control panel in Rows 1-4)
    ws.Rows("5:" & ws.Rows.Count).ClearContents
    
    ' Import starting at A6 (Headers on A5)
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A6"))
        .TextFileParseType = xlDelimited
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2, 2)
        .Refresh BackgroundQuery:=False
    End With
    
    ' Label Headers on Row 5
    ws.Range("A5:B5").Value = Array("RAW_ROW_DATA", "RAW_SPEC_INS")
    ws.UsedRange.Columns.AutoFit
End Sub
