Attribute VB_Name = "PublicReportPublisher"
Option Explicit

' ============================================================
' Public Report Publisher & Discord Uploader
' ------------------------------------------------------------
' Part of the Clan Operations & Analytics Automation Tool.
'
' This module generates a sanitized, public-facing version of
' the internal war log workbook, saves it to a shared cloud
' folder (e.g., OneDrive), and optionally uploads it to a
' Discord channel via webhook.
'
' ENGINEERING TRADE-OFFS:
' 1. File Uploads: Native VBA HTTP objects do not support
'    multipart/form-data file uploads without complex MIME
'    boundary construction. This module bypasses that by
'    dynamically generating a batch script to execute curl.exe.
' 2. JSON Parsing: To avoid external dependencies, the Discord
'    message ID is extracted using targeted string manipulation
'    rather than a full JSON parser.
' ============================================================

' --- CONFIGURATION ---
' Master (Internal) Sheet Names
Private Const MASTER_LOG_SHEET As String = "CurrentCwl"
Private Const MASTER_ANALYSIS_SHEET As String = "Analysis"

' Public (External) Sheet Names
Private Const TARGET_LOG_SHEET As String = "CWL"
Private Const TARGET_ANALYSIS_SHEET As String = "Analysis"
Private Const PUBLIC_FILE_NAME As String = "Public_War_Log.xlsx"

' File Paths
' NOTE: Update this path to match your local OneDrive/SharePoint sync folder.
Private Const PUBLIC_FILE_PATH As String = "C:\Your\OneDrive\Path\Public_War_Log.xlsx"

' Discord Webhook
' WARNING: Never commit a real webhook URL to a public repository.
Private Const DISCORD_WEBHOOK_URL As String = "YOUR_DISCORD_WEBHOOK_URL_HERE"
' ---------------------

Public Sub PublishToDiscord()

    Dim wsMasterLog As Worksheet
    Dim wsMasterAnalysis As Worksheet
    Dim wbNew As Workbook
    Dim wsLog As Worksheet
    Dim wsAnalysis As Worksheet
    Dim tempPath As String
    Dim wb As Workbook
    
    Dim oldScreen As Boolean
    Dim oldAlerts As Boolean
    Dim oldCalc As XlCalculation
    Dim createdNewFile As Boolean
    Dim uploadOK As Boolean

    ' 1. Basic validation
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Please save this workbook before publishing.", vbCritical, "Error"
        Exit Sub
    End If

    If Not SheetExists(ThisWorkbook, MASTER_LOG_SHEET) Or _
       Not SheetExists(ThisWorkbook, MASTER_ANALYSIS_SHEET) Then
        MsgBox "Required master sheets not found.", vbCritical, "Error"
        Exit Sub
    End If

    Set wsMasterLog = ThisWorkbook.Worksheets(MASTER_LOG_SHEET)
    Set wsMasterAnalysis = ThisWorkbook.Worksheets(MASTER_ANALYSIS_SHEET)

    tempPath = Environ("Temp") & Application.PathSeparator & _
               "Public_War_Log_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"

    ' 2. Freeze Excel environment for performance
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    ' 3. Close the public file if it is open in this Excel session
    For Each wb In Workbooks
        If StrComp(wb.FullName, PUBLIC_FILE_PATH, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb

    ' 4. Build a clean temporary workbook
    wsMasterLog.Copy
    Set wbNew = ActiveWorkbook

    Set wsLog = wbNew.Sheets(1)
    wsLog.Name = TARGET_LOG_SHEET
    wsLog.Visible = xlSheetVisible
    
    ' Pull specific columns as values to prevent broken external references
    wsMasterLog.Range("Q:R").Copy
    wsLog.Range("Q1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Remove tables to prevent duplicate table-name conflicts
    Do While wsLog.ListObjects.Count > 0
        wsLog.ListObjects(1).Unlist
    Loop

    wsMasterAnalysis.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
    Set wsAnalysis = wbNew.Sheets(wbNew.Sheets.Count)
    wsAnalysis.Name = TARGET_ANALYSIS_SHEET
    wsAnalysis.Visible = xlSheetVisible

    ' 5. Clean and redact the public sheets
    CleanWarLogSheet wsLog
    CleanAnalysisSheet wsAnalysis
    CleanWorkbook wbNew

    ' 6. Save to temporary file
    If Dir(tempPath) <> vbNullString Then Kill tempPath
    wbNew.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

    ' 7. Replace the public file safely
    If Dir(PUBLIC_FILE_PATH) = vbNullString Then
        Name tempPath As PUBLIC_FILE_PATH
        createdNewFile = True
    Else
        If Not OverwriteFile(tempPath, PUBLIC_FILE_PATH) Then
            MsgBox "Could not overwrite " & PUBLIC_FILE_NAME & "." & vbCrLf & _
                   "Ensure it is closed in Excel desktop, Excel Online, and mobile.", _
                   vbCritical, "File Lock Error"
            GoTo CleanExit
        End If
    End If

    ' 8. Optional Discord upload
    If Len(DISCORD_WEBHOOK_URL) > 0 And DISCORD_WEBHOOK_URL <> "YOUR_DISCORD_WEBHOOK_URL_HERE" Then
        uploadOK = UploadToDiscordWithCurl(PUBLIC_FILE_PATH, DISCORD_WEBHOOK_URL)

        If uploadOK Then
            MsgBox "Public file updated and uploaded to Discord.", vbInformation, "Success"
        Else
            MsgBox "Public file updated locally, but Discord upload failed." & vbCrLf & _
                   "Check webhook URL and curl.exe availability.", vbExclamation, "Upload Failed"
        End If
    Else
        If createdNewFile Then
            MsgBox "Created a new " & PUBLIC_FILE_NAME & "." & vbCrLf & _
                   "Please create/share a new public link for it.", vbInformation, "Success"
        Else
            MsgBox "Public file updated locally.", vbInformation, "Success"
        End If
    End If

CleanExit:
    On Error Resume Next
    If Len(tempPath) > 0 And Dir(tempPath) <> vbNullString Then Kill tempPath
    
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    Exit Sub

ErrHandler:
    MsgBox "Publish failed: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
    If Not wbNew Is Nothing Then
        On Error Resume Next
        wbNew.Close SaveChanges:=False
    End If
    Resume CleanExit
End Sub

' ============================================================
' Helper: Sheet Validation
' ============================================================
Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' ============================================================
' Helper: Safe File Overwrite (with retry logic for cloud sync locks)
' ============================================================
Private Function OverwriteFile(ByVal sourcePath As String, ByVal destPath As String) As Boolean
    Dim attempt As Long
    Dim errNum As Long

    ' Retry up to 5 times because OneDrive/Antivirus can briefly lock files
    For attempt = 1 To 5
        On Error Resume Next
        SetAttr destPath, vbNormal
        Err.Clear
        FileCopy sourcePath, destPath
        errNum = Err.Number
        On Error GoTo 0

        If errNum = 0 Then
            OverwriteFile = True
            Exit Function
        End If

        If attempt < 5 Then Application.Wait Now + TimeValue("00:00:02")
    Next attempt

    OverwriteFile = False
End Function

' ============================================================
' Helper: War Log Sanitization
' ============================================================
Private Sub CleanWarLogSheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long

    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
    End If

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Unlist
    Loop

    lastRow = GetLastRow(ws)
    For i = lastRow To 1 Step -1
        If ws.Rows(i).Hidden Then ws.Rows(i).Delete
    Next i

    ConvertFormulasToValues ws

    ' Delete unwanted internal columns (Right to Left to preserve indexes)
    On Error Resume Next
    ws.Columns("S:" & ws.Columns.Count).Delete ' Delete S to end
    ws.Columns("K:P").Delete
    ws.Columns("D:G").Delete
    On Error GoTo 0
End Sub

' ============================================================
' Helper: Analysis Sheet Sanitization
' ============================================================
Private Sub CleanAnalysisSheet(ByVal ws As Worksheet)
    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
    End If

    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode = False
    On Error GoTo 0

    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Unlist
    Loop

    ConvertFormulasToValues ws

    On Error Resume Next
    ws.Columns("L:" & ws.Columns.Count).Delete
    On Error GoTo 0
End Sub

' ============================================================
' Helper: Workbook Cleanup (Remove named ranges and connections)
' ============================================================
Private Sub CleanWorkbook(ByVal wb As Workbook)
    Dim n As Long
    On Error Resume Next
    For n = wb.Names.Count To 1 Step -1
        wb.Names(n).Delete
    Next n
    For n = wb.Connections.Count To 1 Step -1
        wb.Connections(n).Delete
    Next n
    On Error GoTo 0
End Sub

' ============================================================
' Helper: Range Utilities
' ============================================================
Private Function GetLastRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not c Is Nothing Then GetLastRow = c.Row
End Function

Private Sub ConvertFormulasToValues(ByVal ws As Worksheet)
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Sub
    
    ur.Value = ur.Value
    If Err.Number <> 0 Then
        Err.Clear
        ur.Copy
        ur.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    On Error GoTo 0
End Sub

' ============================================================
' Discord Integration (via curl.exe)
' ============================================================
Private Function UploadToDiscordWithCurl(ByVal filePath As String, ByVal webhookUrl As String) As Boolean
    Dim fso As Object, wsh As Object
    Dim tmpFolder As String, stamp As String
    Dim patchBatPath As String, postBatPath As String
    Dim payloadJsonPath As String, responseBodyPath As String, httpCodePath As String
    Dim cmd As String, exitCode As Long
    Dim httpCode As String, responseText As String
    Dim fNum As Integer
    Dim msgIdFilePath As String, lastMsgId As String, newMsgId As String
    Dim baseUrl As String, editUrl As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")

    tmpFolder = Environ("TEMP")
    If Right(tmpFolder, 1) <> Application.PathSeparator Then tmpFolder = tmpFolder & Application.PathSeparator
    stamp = Format(Now, "yyyymmddhhnnss")

    patchBatPath = tmpFolder & "discord_patch_" & stamp & ".bat"
    postBatPath = tmpFolder & "discord_post_" & stamp & ".bat"
    payloadJsonPath = tmpFolder & "payload_" & stamp & ".json"
    responseBodyPath = tmpFolder & "discord_response_" & stamp & ".txt"
    httpCodePath = tmpFolder & "discord_http_" & stamp & ".txt"

    ' Store message ID in AppData to allow editing previous messages
    msgIdFilePath = Environ("APPDATA") & "\CWL_Discord"
    If Not fso.FolderExists(msgIdFilePath) Then fso.CreateFolder msgIdFilePath
    msgIdFilePath = msgIdFilePath & "\last_msg_id.txt"

    baseUrl = Trim(webhookUrl)
    baseUrl = Replace(Replace(baseUrl, "?wait=true", ""), "&wait=true", "")
    If Right(baseUrl, 1) = "/" Then baseUrl = Left(baseUrl, Len(baseUrl) - 1)

    lastMsgId = Trim(ReadTextFile(msgIdFilePath))

    ' STEP 1: Try to EDIT (PATCH) the existing message to avoid channel spam
    If Len(lastMsgId) > 5 Then
        editUrl = baseUrl & "/messages/" & lastMsgId

        fNum = FreeFile
        Open payloadJsonPath For Output As #fNum
        Print #fNum, "{""attachments"":[]}" ' Clear old attachment
        Close #fNum

        fNum = FreeFile
        Open patchBatPath For Output As #fNum
        Print #fNum, "@echo off"
        Print #fNum, "chcp 65001 >nul"
        Print #fNum, "curl.exe -s -S -X PATCH -F ""files[0]=@" & filePath & """ -F ""payload_json=<" & payloadJsonPath & """ -o """ & responseBodyPath & """ -w ""%%{http_code}"" """ & editUrl & """ > """ & httpCodePath & """"
        Close #fNum

        exitCode = wsh.Run("cmd.exe /c """ & patchBatPath & """", 0, True)
        httpCode = Trim(ReadTextFile(httpCodePath))

        On Error Resume Next
        fso.DeleteFile patchBatPath, True
        fso.DeleteFile payloadJsonPath, True
        On Error GoTo 0

        If exitCode = 0 And httpCode Like "2*" Then
            UploadToDiscordWithCurl = True
            On Error Resume Next
            fso.DeleteFile responseBodyPath, True
            fso.DeleteFile httpCodePath, True
            On Error GoTo 0
            Exit Function
        End If
    End If

    ' STEP 2: POST a new message (if PATCH failed or first run)
    fNum = FreeFile
    Open postBatPath For Output As #fNum
    Print #fNum, "@echo off"
    Print #fNum, "chcp 65001 >nul"
    Print #fNum, "curl.exe -s -S -F ""files[0]=@" & filePath & """ -o """ & responseBodyPath & """ -w ""%%{http_code}"" """ & baseUrl & "?wait=true"" > """ & httpCodePath & """"
    Close #fNum

    exitCode = wsh.Run("cmd.exe /c """ & postBatPath & """", 0, True)
    httpCode = Trim(ReadTextFile(httpCodePath))
    responseText = Trim(ReadTextFile(responseBodyPath))

    On Error Resume Next
    fso.DeleteFile postBatPath, True
    fso.DeleteFile responseBodyPath, True
    fso.DeleteFile httpCodePath, True
    On Error GoTo 0

    If exitCode = 0 And (httpCode Like "2*" Or httpCode = "") Then
        UploadToDiscordWithCurl = True
        newMsgId = ExtractMessageId(responseText)
        If Len(newMsgId) > 5 Then SaveTextFile msgIdFilePath, newMsgId
    Else
        UploadToDiscordWithCurl = False
    End If
End Function

' ============================================================
' Helper: Manual JSON Parsing for Message ID
' ============================================================
Private Function ExtractMessageId(ByVal jsonResponse As String) As String
    Dim channelPos As Long, idKeyPos As Long, colonPos As Long
    Dim quote1 As Long, quote2 As Long, candidate As String
    
    ' Find the real message ID by locating "channel_id" and finding 
    ' the last "id" key that appears before it (ignoring attachment IDs).
    channelPos = InStr(1, jsonResponse, """channel_id""", vbTextCompare)
    If channelPos = 0 Then Exit Function
    
    idKeyPos = InStrRev(Left$(jsonResponse, channelPos - 1), """id""")
    If idKeyPos = 0 Then Exit Function
    
    colonPos = InStr(idKeyPos, jsonResponse, ":")
    If colonPos = 0 Then Exit Function
    
    quote1 = InStr(colonPos + 1, jsonResponse, """")
    If quote1 = 0 Then Exit Function
    
    quote2 = InStr(quote1 + 1, jsonResponse, """")
    If quote2 = 0 Then Exit Function
    
    candidate = Mid$(jsonResponse, quote1 + 1, quote2 - quote1 - 1)
    If IsAllDigits(candidate) Then ExtractMessageId = candidate
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

' ============================================================
' Helper: File I/O
' ============================================================
Private Sub SaveTextFile(ByVal filePath As String, ByVal content As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True)
    ts.Write content
    ts.Close
End Sub

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then ReadTextFile = fso.OpenTextFile(filePath, 1).ReadAll
End Sub

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(folderPath)
End Function