Attribute VB_Name = "DatabaseSync"
Option Explicit

' ============================================================
' Database Synchronization Module
' ------------------------------------------------------------
' Part of the Clan Operations & Analytics Automation Tool.
'
' This module performs a simple ETL (Extract, Transform, Load)
' operation. It reads updated metrics from the Analysis sheet
' and synchronizes them with the master Database sheet based
' on a primary key (Player Name).
' ============================================================

' --- CONFIGURATION ---
Private Const SHEET_ANALYSIS As String = "Analysis"
Private Const SHEET_DATABASE As String = "Database"

' Column Indexes (1 = A, 4 = D, 12 = L)
Private Const COL_PLAYER_NAME As Long = 1
Private Const COL_NEW_POWER As Long = 12
Private Const COL_DB_POWER As Long = 4
' ---------------------

Public Sub SyncNewPower()
    Dim wsAnalysis As Worksheet
    Dim wsDatabase As Worksheet
    Dim dbSearchRange As Range
    
    Dim i As Long
    Dim lastRowAnalysis As Long
    Dim lastRowDb As Long
    
    Dim playerName As String
    Dim newPower As Variant
    Dim foundCell As Range
    
    Set wsAnalysis = ThisWorkbook.Sheets(SHEET_ANALYSIS)
    Set wsDatabase = ThisWorkbook.Sheets(SHEET_DATABASE)
    
    ' Find the bottom of the Analysis list
    lastRowAnalysis = wsAnalysis.Cells(wsAnalysis.Rows.Count, COL_PLAYER_NAME).End(xlUp).Row
    
    ' Define a dynamic search range in the Database instead of a hardcoded A2:A200
    lastRowDb = wsDatabase.Cells(wsDatabase.Rows.Count, COL_PLAYER_NAME).End(xlUp).Row
    If lastRowDb < 2 Then lastRowDb = 2 ' Ensure range is valid even if DB is empty
    Set dbSearchRange = wsDatabase.Range(wsDatabase.Cells(2, COL_PLAYER_NAME), wsDatabase.Cells(lastRowDb, COL_PLAYER_NAME))
    
    ' Loop through the Analysis sheet and update the Database
    For i = 2 To lastRowAnalysis
        playerName = Trim(CStr(wsAnalysis.Cells(i, COL_PLAYER_NAME).Value))
        newPower = wsAnalysis.Cells(i, COL_NEW_POWER).Value
        
        ' Only process if there is a valid player name and a new power value
        If Len(playerName) > 0 And Len(Trim(CStr(newPower))) > 0 Then
            
            Set foundCell = dbSearchRange.Find(What:=playerName, LookAt:=xlWhole, MatchCase:=False)
            
            If Not foundCell Is Nothing Then
                ' Update the database only if the value has actually changed
                If wsDatabase.Cells(foundCell.Row, COL_DB_POWER).Value <> newPower Then
                    wsDatabase.Cells(foundCell.Row, COL_DB_POWER).Value = newPower
                End If
            End If
        End If
    Next i

    MsgBox "Database synchronization complete.", vbInformation, "Sync Complete"
End Sub