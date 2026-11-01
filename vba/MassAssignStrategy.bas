Attribute VB_Name = "CwlStrategyAssignment"
Option Explicit

' ============================================================
' CWL Default Strategy Assignment
' ------------------------------------------------------------
' Part of the Clan Operations & Analytics Automation Tool.
'
' This module applies a default star strategy (2-star or 3-star)
' to all active roster members for a specific CWL day who do
' not already have a target assignment.
' ============================================================

' Sheet and Cell Configuration
Private Const SHEET_NAME As String = "CurrentCwl"
Private Const CELL_CURRENT_DAY As String = "N2"

' Column Layout
Private Const COL_DAY As Long = 1             ' A
Private Const COL_PLAYER_NAME As Long = 3     ' C
Private Const COL_ASSIGNED_TARGET As Long = 8 ' H
Private Const COL_STRATEGY As Long = 9        ' I

Public Sub AssignDefaultCwlStrategy()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)

    ' 1. Verify the current day from the configuration cell
    Dim currentDay As Long
    If Not IsNumeric(ws.Range(CELL_CURRENT_DAY).Value) Or ws.Range(CELL_CURRENT_DAY).Value = "" Then
        MsgBox "Cell " & CELL_CURRENT_DAY & " does not contain a valid day number.", vbCritical, "Configuration Error"
        Exit Sub
    End If
    currentDay = CLng(ws.Range(CELL_CURRENT_DAY).Value)

    ' 2. Prompt the user to select the default strategy
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("Select the default strategy for Day " & currentDay & ":" & vbCrLf & vbCrLf & _
                        "[YES] = Assign 3 Stars" & vbCrLf & _
                        "[NO] = Assign 2 Stars", _
                        vbYesNoCancel + vbQuestion, "Strategy Assignment")

    Dim strategyStars As String
    Dim starSymbol As String
    starSymbol = ChrW(&H2B50) ' Unicode star character

    If userChoice = vbYes Then
        strategyStars = starSymbol & starSymbol & starSymbol
    ElseIf userChoice = vbNo Then
        strategyStars = starSymbol & starSymbol
    Else
        Exit Sub ' User clicked Cancel or closed the dialog
    End If

    ' 3. Scan the roster and apply the strategy
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_PLAYER_NAME).End(xlUp).Row

    Dim i As Long
    Dim assignedCount As Long
    Dim skippedCount As Long

    assignedCount = 0
    skippedCount = 0

    ' Disable screen updating for faster execution during batch writes
    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    For i = 2 To lastRow
        ' Check if the row belongs to the current day and has a player name
        If Val(ws.Cells(i, COL_DAY).Value) = currentDay And Len(Trim$(CStr(ws.Cells(i, COL_PLAYER_NAME).Value))) > 0 Then

            ' Only assign a strategy if the player does not already have a target assignment
            If Len(Trim$(CStr(ws.Cells(i, COL_ASSIGNED_TARGET).Value))) = 0 Then
                ws.Cells(i, COL_STRATEGY).Value = strategyStars
                assignedCount = assignedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If

        End If
    Next i

    ' Restore screen updating
    Application.ScreenUpdating = oldScreenUpdating

    ' 4. Final status report
    MsgBox "Strategy assigned for " & assignedCount & " players on Day " & currentDay & "." & vbCrLf & _
           "Skipped " & skippedCount & " players with pre-existing assignments.", _
           vbInformation, "Assignment Complete"

End Sub