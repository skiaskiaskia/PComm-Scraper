Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub RunMainframeScrape()
    Dim objShell As Object, fso As Object
    Dim vbsPath As String, csvPath As String, logPath As String, vbsCode As String
    Dim sName As String ' <--- This was the missing variable!
    
    vbsPath = "C:\temp\PCommBridge.vbs"
    csvPath = "C:\temp\mainframe_data.csv"
    logPath = "C:\temp\trace.log"
    sName = "Unknown" ' Default value
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")

    ' Ensure C:\temp exists
    If Not fso.FolderExists("C:\temp") Then
        MsgBox "The folder C:\temp does not exist. Please create it first.", vbCritical
        Exit Sub
    End If

    If fso.FileExists(csvPath) Then fso.DeleteFile csvPath
    If fso.FileExists(logPath) Then fso.DeleteFile logPath

    ' --- BUILD THE VBSCRIPT ---
    vbsCode = "Option Explicit" & vbCrLf
    vbsCode = vbsCode & "Dim objSess, objPS, objOIA, objFSO, objFile, objLog, r, txt, buf, rows, cols, clist, i, sName" & vbCrLf
    vbsCode = vbsCode & "Set objFSO = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf
    vbsCode = vbsCode & "Sub UpdateLog(msg) : Set objLog = objFSO.CreateTextFile(""" & logPath & """, True) : objLog.WriteLine(msg) : objLog.Close : End Sub" & vbCrLf
    
    vbsCode = vbsCode & "UpdateLog ""START: Initializing File System""" & vbCrLf
    vbsCode = vbsCode & "Set objFile = objFSO.CreateTextFile(""" & csvPath & """, True)" & vbCrLf
    vbsCode = vbsCode & "UpdateLog ""STEP: Searching for PComm Sessions""" & vbCrLf
    
    vbsCode = vbsCode & "Set clist = CreateObject(""PCOMM.autECLConnList"")" & vbCrLf
    vbsCode = vbsCode & "clist.Refresh" & vbCrLf
    vbsCode = vbsCode & "If clist.Count = 0 Then UpdateLog ""ERROR: No PComm windows found"": MsgBox ""No PComm Found"": WScript.Quit" & vbCrLf
    
    vbsCode = vbsCode & "sName = """" " & vbCrLf
    vbsCode = vbsCode & "For i = 1 To clist.Count" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, clist(i).Name, ""BTS Pooled"", 1) > 0 Then sName = clist(i).Name : Exit For" & vbCrLf
    vbsCode = vbsCode & "Next" & vbCrLf
    
    vbsCode = vbsCode & "If sName = """" Then sName = clist(1).Name" & vbCrLf
    vbsCode = vbsCode & "UpdateLog ""SESSION:"" & sName" & vbCrLf ' We tag the session name in the log
    
    vbsCode = vbsCode & "On Error Resume Next" & vbCrLf
    vbsCode = vbsCode & "Set objSess = CreateObject(""PCOMM.autECLSession"")" & vbCrLf
    vbsCode = vbsCode & "objSess.SetConnectionByName(sName)" & vbCrLf
    
    vbsCode = vbsCode & "Set objPS = objSess.autECLPS : Set objOIA = objSess.autECLOIA" & vbCrLf
    vbsCode = vbsCode & "If Err.Number <> 0 Then UpdateLog ""ERROR: Failed to bind to Session Objects"": WScript.Quit" & vbCrLf
    vbsCode = vbsCode & "rows = objPS.NumRows : cols = objPS.NumCols" & vbCrLf
    
    vbsCode = vbsCode & "Do" & vbCrLf
    vbsCode = vbsCode & "  objOIA.WaitForInputReady" & vbCrLf
    vbsCode = vbsCode & "  For r = 5 To (rows - 3)" & vbCrLf
    vbsCode = vbsCode & "    txt = Trim(objPS.GetText(r, 1, cols))" & vbCrLf
    vbsCode = vbsCode & "    objFile.WriteLine Chr(34) & txt & Chr(34)" & vbCrLf
    vbsCode = vbsCode & "  Next" & vbCrLf
    
    vbsCode = vbsCode & "  buf = objPS.GetText(1, 1, rows * cols)" & vbCrLf
    vbsCode = vbsCode & "  If InStr(1, buf, ""invalid page forward"", 1) > 0 Or InStr(1, buf, ""invalid"", 1) > 0 Then Exit Do" & vbCrLf
    vbsCode = vbsCode & "  UpdateLog ""LOOP: Sending PA1 key""" & vbCrLf
    vbsCode = vbsCode & "  objPS.SendKeys ""[pa1]""" & vbCrLf
    vbsCode = vbsCode & "  WScript.Sleep 800" & vbCrLf
    vbsCode = vbsCode & "Loop" & vbCrLf
    
    vbsCode = vbsCode & "UpdateLog ""FINISH: Data extraction complete""" & vbCrLf
    vbsCode = vbsCode & "objFile.Close"

    ' 3. WRITE AND EXECUTE
    fso.CreateTextFile(vbsPath, True).Write vbsCode
    
    On Error Resume Next
    objShell.Run "C:\Windows\SysWOW64\wscript.exe """ & vbsPath & """", 1, True
    
    ' 4. RETRIEVE SESSION NAME FROM LOG FOR THE FINAL MESSAGE
    If fso.FileExists(logPath) Then
        Dim ts As Object, line As String
        Set ts = fso.OpenTextFile(logPath, 1)
        Do Until ts.AtEndOfStream
            line = ts.ReadLine
            If Left(line, 8) = "SESSION:" Then sName = Mid(line, 9)
        Loop
        ts.Close
    End If

    ' 5. CHECK FAILURES
    If Not fso.FileExists(csvPath) Then
        Dim lastStatus As String
        If fso.FileExists(logPath) Then
            lastStatus = fso.OpenTextFile(logPath).ReadLine
        Else
            lastStatus = "VBScript failed to even start."
        End If
        MsgBox "SCRAPE FAILED!" & vbCrLf & "Last status: " & lastStatus, vbCritical
        Exit Sub
    End If

    ' 6. IMPORT DATA
    ImportData csvPath
    
    ' Cleanup
    If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath
    MsgBox "Success! Data loaded from Session: " & sName, vbInformation
End Sub

Sub ImportData(path As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Columns("B").ClearContents
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("B2"))
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileColumnDataTypes = Array(2)
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.UsedRange.SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub
