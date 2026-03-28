' --- BUILD THE VBSCRIPT (TAB-SENSING VERSION) ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, r, txt, detail, rows, cols, clist, i, sName, regEx, lastRow" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Set regEx = CreateObject(""VBScript.RegExp"")" & vbCrLf
    vbsCode = vbsCode & "regEx.Global = True : regEx.Pattern = ""\s{2,}""" & vbCrLf
    
    ' ... (Session Connection Logic remains same) ...

    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    ' Move to just before the first data row to start tabbing
    vbsCode = vbsCode & "  objPS.SetCursorPos 7, 80" & vbCrLf 
    vbsCode = vbsCode & "  lastRow = 0" & vbCrLf
    
    vbsCode = vbsCode & "  Do" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[tab]""" & vbCrLf
    vbsCode = vbsCode & "    WScript.Sleep 50" & vbCrLf
    
    ' Check if we are still in the data zone (Rows 8 to 22)
    ' If the cursor jumps to Row 23+ or back to Row 1-7, we stop tabbing this page
    vbsCode = vbsCode & "    If objPS.CursorRow < 8 Or objPS.CursorRow > 22 Then Exit Do" & vbCrLf
    
    ' Prevent infinite loops on the same row
    vbsCode = vbsCode & "    If objPS.CursorRow <= lastRow Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "    lastRow = objPS.CursorRow" & vbCrLf
    vbsCode = vbsCode & "    r = objPS.CursorRow" & vbCrLf
    
    ' --- SCRAPE & DRILL ---
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf2]""" & vbCrLf ' Drill down
    vbsCode = vbsCode & "    WScript.Sleep 500" & vbCrLf
    
    ' GRAB DETAIL (Update 5, 10, 20 to your actual detail location)
    vbsCode = vbsCode & "    detail = Trim(objPS.GetText(5, 10, 20))" & vbCrLf 
    
    vbsCode = vbsCode & "    objPS.SendKeys ""[pf3]""" & vbCrLf ' Return
    vbsCode = vbsCode & "    WScript.Sleep 500" & vbCrLf
    
    ' Format and save
    vbsCode = vbsCode & "    txt = regEx.Replace(txt, ""|"")" & vbCrLf
    vbsCode = vbsCode & "    objFile.WriteLine txt & ""|"" & detail" & vbCrLf
    vbsCode = vbsCode & "  Loop" & vbCrLf ' End of Tab Loop
    
    ' Advance Page
    vbsCode = vbsCode & "  If InStr(1, UCase(objPS.GetText(1, 1, rows*cols)), ""LAST PAGE"") > 0 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep 800" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
