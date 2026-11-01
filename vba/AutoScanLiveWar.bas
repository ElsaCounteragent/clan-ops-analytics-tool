Attribute VB_Name = "WarApiTracker"
Option Explicit

' ============================================================
' War API Tracker & Performance Logger
' ------------------------------------------------------------
' Part of the Clan Operations & Analytics Automation Tool.
'
' This module queries the Clash of Clans API for the current
' war state, parses the response, and tracks member performance
' based on attack targets and star counts.
'
' NOTE: This uses native string splitting to parse JSON to avoid
' external library dependencies. For a production environment,
' a dedicated JSON parser (like VBA-JSON) is recommended.
' ============================================================

' --- CONFIGURATION ---
' WARNING: Never commit a real API token to a public repository.
' Use a placeholder here and store the real token in a hidden sheet,
' a local text file, or prompt the user at runtime.
Private Const API_TOKEN As String = "YOUR_API_TOKEN_HERE"
Private Const CLAN_TAG As String = "%23CLANTAG" ' URL-encoded #CLANTAG
Private Const API_BASE_URL As String = "https://api.clashofclans.com/v1/clans/"

Private Const SHEET_TRACKER As String = "WarTracker"
Private Const SHEET_DATABASE As String = "Database"
Private Const PROMOTION_THRESHOLD As Long = 3
' ---------------------

Public Sub AutoScanLiveWar()
    Dim http As Object
    Dim jsonResponse As String
    Dim clanData As String, opponentData As String
    Dim url As String
    Dim wsTracker As Worksheet, wsDatabase As Worksheet
    Dim prepTime As String
    
    Set wsTracker = ThisWorkbook.Sheets(SHEET_TRACKER)
    Set wsDatabase = ThisWorkbook.Sheets(SHEET_DATABASE)
    url = API_BASE_URL & CLAN_TAG & "/currentwar"
    
    On Error GoTo SafeExit
    
    ' 1. Query the API
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & API_TOKEN
    http.setRequestHeader "Accept", "application/json"
    http.send
    jsonResponse = http.responseText
    
    ' 2. Guard Rails: Only execute during live war or war ended states
    If InStr(jsonResponse, "reason") > 0 Or jsonResponse = "" Then GoTo SafeExit
    If InStr(jsonResponse, """state"":""inWar""") = 0 And InStr(jsonResponse, """state"":""warEnded""") = 0 Then GoTo SafeExit
    
    ' 3. Extract the Unique War Timestamp (used as a unique war ID)
    prepTime = Split(Split(jsonResponse, """preparationStartTime"":""")(1), """,")(0)
    
    ' 4. Split the payload into Clan vs Opponent blocks
    clanData = Split(jsonResponse, """opponent"":{")(0)
    opponentData = Split(jsonResponse, """opponent"":{")(1)
    
    ' 5. Build a Dictionary of Opponent Map Positions
    Dim oppDict As Object
    Set oppDict = CreateObject("Scripting.Dictionary")
    
    Dim oppBlocks() As String
    Dim oppTag As String, oppMap As Long
    Dim i As Long
    
    oppBlocks = Split(opponentData, "{""tag"":""")
    For i = 1 To UBound(oppBlocks)
        oppTag = Split(oppBlocks(i), """")(0)
        If InStr(oppBlocks(i), """mapPosition"":") > 0 Then
            oppMap = Val(Split(Split(oppBlocks(i), """mapPosition"":")(1), ",")(0))
            If Not oppDict.Exists(oppTag) Then
                oppDict.Add oppTag, oppMap
            End If
        End If
    Next i
    
    ' 6. Evaluate Clan Roster Performance
    Dim clanBlocks() As String
    Dim playerName As String, attackerMap As Long
    Dim attacksBlock As String, attackChunks() As String
    Dim stars1 As Long, stars2 As Long
    Dim defTag1 As String, defTag2 As String
    Dim passedTrial As Boolean
    
    clanBlocks = Split(clanData, "{""tag"":""")
    
    For i = 1 To UBound(clanBlocks)
        passedTrial = False
        
        If InStr(clanBlocks(i), """mapPosition"":") > 0 Then
            playerName = Split(Split(clanBlocks(i), """name"":""")(1), """")(0)
            attackerMap = Val(Split(Split(clanBlocks(i), """mapPosition"":")(1), ",")(0))
            
            ' Process attack data
            If InStr(clanBlocks(i), """attacks"":[") > 0 Then
                attacksBlock = Split(clanBlocks(i), """attacks"":[")(1)
                attacksBlock = Split(attacksBlock, "]")(0)
                attackChunks = Split(attacksBlock, "{""attackerTag"":")
                
                ' Check if exactly 2 attacks exist
                If UBound(attackChunks) = 2 Then
                    stars1 = Val(Split(Split(attackChunks(1), """stars"":")(1), ",")(0))
                    defTag1 = Split(Split(attackChunks(1), """defenderTag"":""")(1), """")(0)
                    
                    stars2 = Val(Split(Split(attackChunks(2), """stars"":")(1), ",")(0))
                    defTag2 = Split(Split(attackChunks(2), """defenderTag"":""")(1), """")(0)
                    
                    ' Validation: 6 Total Stars Requirement
                    If stars1 = 3 And stars2 = 3 Then
                        If oppDict.Exists(defTag1) And oppDict.Exists(defTag2) Then
                            
                            ' Map Position 1 (Leadership) Exception:
                            ' The #1 player must attack the enemy #1 and any other target.
                            If attackerMap = 1 Then
                                If oppDict(defTag1) = 1 Or oppDict(defTag2) = 1 Then
                                    passedTrial = True
                                End If
                            Else
                                ' Standard Rule:
                                ' All other players must hit their equal or higher map position for both attacks.
                                If oppDict(defTag1) <= attackerMap And oppDict(defTag2) <= attackerMap Then
                                    passedTrial = True
                                End If
                            End If
                            
                        End If
                    End If
                End If
            End If
            
            ' 7. Logging & Promotion Logic
            If passedTrial Then
                Call LogPlayerPerformance(wsTracker, wsDatabase, playerName, prepTime)
            End If
        End If
    Next i

SafeExit:
    Set http = Nothing
    Set oppDict = Nothing
End Sub

' ============================================================
' Helper: Logging & Promotion
' ============================================================
Private Sub LogPlayerPerformance(wsTracker As Worksheet, wsDatabase As Worksheet, _
                               ByVal playerName As String, ByVal prepTime As String)
                               
    Dim foundCell As Range
    Dim nextRow As Long
    
    Set foundCell = wsTracker.Range("A:A").Find(What:=playerName, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        ' New entry to tracker
        nextRow = wsTracker.Cells(wsTracker.Rows.Count, "A").End(xlUp).Row + 1
        wsTracker.Cells(nextRow, 1).Value = playerName
        wsTracker.Cells(nextRow, 2).Value = 1          ' Success count
        wsTracker.Cells(nextRow, 3).Value = prepTime   ' Last war timestamp
    Else
        ' Check individual war lock (prevent double counting the same war)
        If foundCell.Offset(0, 2).Value <> prepTime Then
            foundCell.Offset(0, 1).Value = foundCell.Offset(0, 1).Value + 1
            foundCell.Offset(0, 2).Value = prepTime
            
            ' Handle main Database promotion
            If foundCell.Offset(0, 1).Value >= PROMOTION_THRESHOLD Then
                Call PromoteToDatabase(wsDatabase, playerName)
            End If
        End If
    End If
End Sub

Private Sub PromoteToDatabase(wsDatabase As Worksheet, ByVal playerName As String)
    Dim dbCell As Range
    Dim dbNextRow As Long
    
    Set dbCell = wsDatabase.Range("A:A").Find(What:=playerName, LookAt:=xlWhole)
    If dbCell Is Nothing Then
        dbNextRow = wsDatabase.Cells(wsDatabase.Rows.Count, "A").End(xlUp).Row + 1
        wsDatabase.Cells(dbNextRow, 1).Value = playerName
    End If
End Sub