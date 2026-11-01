Attribute VB_Name = "DiscordNotifications"
Option Explicit

' ============================================================
' Discord Notification Module
' ------------------------------------------------------------
' Part of the Clan Operations & Analytics Automation Tool.
'
' This module handles batched Discord webhook notifications.
' It reads assignment messages from a worksheet range, sends
' them sequentially with a delay to respect API rate limits,
' and schedules a follow-up role mention.
' ============================================================

' --- CONFIGURATION ---
' WARNING: Never commit real webhook URLs or Role IDs to a public repository.
Private Const WEBHOOK_URL_ASSIGNMENTS As String = "YOUR_DISCORD_WEBHOOK_URL_HERE"
Private Const WEBHOOK_URL_ANNOUNCEMENT As String = "YOUR_DISCORD_WEBHOOK_URL_HERE"
Private Const DISCORD_ROLE_ID As String = "YOUR_DISCORD_ROLE_ID_HERE"

' Worksheet and Range Configuration
Private Const SHEET_NAME As String = "CurrentCwl"
Private Const ASSIGNMENTS_RANGE As String = "M3:M22"
Private Const DELAY_SECONDS As Long = 1
Private Const FOLLOW_UP_DELAY_MINUTES As Long = 3
' ---------------------

Public Sub SendBatchedAssignments()
    Dim http As Object
    Dim jsonPayload As String
    Dim cell As Range
    Dim commsRange As Range
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Set commsRange = ws.Range(ASSIGNMENTS_RANGE)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' 1. Transmit individual assignment messages
    For Each cell In commsRange
        If Len(Trim(CStr(cell.Value))) > 0 Then
            ' Basic JSON formatting (escapes double quotes to prevent broken payloads)
            jsonPayload = "{""content"": """ & Replace(CStr(cell.Value), """", "\""") & """}"
            
            On Error Resume Next
            http.Open "POST", WEBHOOK_URL_ASSIGNMENTS, False
            http.setRequestHeader "Content-Type", "application/json"
            http.Send jsonPayload
            On Error GoTo 0
            
            ' Pause to respect Discord API rate limits
            Application.Wait (Now + TimeValue("0:00:0" & DELAY_SECONDS))
            DoEvents
        End If
    Next cell
    
    ' 2. Schedule the follow-up role ping
    Application.OnTime Now + TimeValue("00:" & Format(FOLLOW_UP_DELAY_MINUTES, "00") & ":00"), "SendRolePing"
    
    MsgBox "Assignment notifications sent." & vbCrLf & vbCrLf & _
           "A follow-up role mention is scheduled in " & FOLLOW_UP_DELAY_MINUTES & " minutes." & vbCrLf & _
           "Please leave this workbook open until the timer completes.", _
           vbInformation, "Discord Sync"
End Sub

' ============================================================
' Scheduled Follow-Up: Role Mention
' ============================================================
Public Sub SendRolePing()
    Dim http As Object
    Dim jsonPayload As String
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Format the Discord role mention syntax: <@&ROLE_ID>
    jsonPayload = "{""content"": ""<@&" & DISCORD_ROLE_ID & "> Assignments for the current event have been posted!""}"
    
    On Error Resume Next
    http.Open "POST", WEBHOOK_URL_ANNOUNCEMENT, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send jsonPayload
    On Error GoTo 0
    
    Set http = Nothing
End Sub