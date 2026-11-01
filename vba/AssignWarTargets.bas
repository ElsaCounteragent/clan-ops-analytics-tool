Attribute VB_Name = "WarAssignments"
Option Explicit

' ============================================================
' War Target Assignment Module
' ------------------------------------------------------------
' Part of the Clan Operations & Analytics Automation Tool.
'
' This module assigns war targets to active players based on:
' - player power
' - target power
' - existing assignments
' - a small risk margin
'
' The tool assists with planning. It does not verify live game
' state or make decisions outside the workbook.
' ============================================================

Private Const DEFAULT_RISK_MARGIN As Double = 0.2

' Worksheet layout
Private Const COL_DAY As Long = 1              ' A
Private Const COL_PLAYER_NAME As Long = 3      ' C
Private Const COL_PLAYER_POWER As Long = 7     ' G
Private Const COL_ASSIGNED_TARGET As Long = 8  ' H
Private Const COL_STRATEGY As Long = 9         ' I
Private Const COL_TARGET_ID As Long = 17       ' Q
Private Const COL_TARGET_POWER As Long = 18    ' R

' ============================================================
' Main entry point
' ============================================================

Public Sub AssignWarTargets()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim riskMargin As Double
    riskMargin = DEFAULT_RISK_MARGIN

    Dim userDay As String
    userDay = InputBox( _
        "Confirm the war day:", _
        "War Target Assignment", _
        ws.Range("N2").Text _
    )

    If Len(Trim$(userDay)) = 0 Then Exit Sub

    Dim currentDay As Long
    currentDay = CLng(Val(userDay))

    If currentDay <= 0 Then
        MsgBox "Invalid war day entered.", vbCritical, "Input Error"
        Exit Sub
    End If

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    ' --------------------------------------------------------
    ' Find the used range based on Column A.
    ' --------------------------------------------------------
    Dim absoluteLast As Long
    absoluteLast = ws.Cells(ws.Rows.Count, COL_DAY).End(xlUp).Row

    If absoluteLast < 2 Then
        MsgBox "The assignment sheet is empty.", vbCritical, "Missing Data"
        GoTo CleanExit
    End If

    Dim Players() As Variant
    Dim Targets() As Variant

    Dim pCount As Long
    Dim tCount As Long
    Dim playerCapacity As Long
    Dim targetCapacity As Long

    Dim i As Long
    Dim j As Long

    Dim assignedRegistry As Object
    Set assignedRegistry = CreateObject("Scripting.Dictionary")
    assignedRegistry.CompareMode = vbTextCompare

    Dim dayFirstRow As Long
    Dim dayLastRow As Long

    dayFirstRow = 0
    dayLastRow = 0

    ' --------------------------------------------------------
    ' Find the first and last row for the selected war day.
    ' Scanning the full range is safer if there are blank rows.
    ' --------------------------------------------------------
    For i = 2 To absoluteLast
        If Val(ws.Cells(i, COL_DAY).Value) = currentDay Then
            If dayFirstRow = 0 Then dayFirstRow = i
            dayLastRow = i
        End If
    Next i

    If dayFirstRow = 0 Then
        MsgBox "No rows found for War Day " & currentDay & ".", vbExclamation, "No Data"
        GoTo CleanExit
    End If

    ' --------------------------------------------------------
    ' Record targets that are already assigned.
    ' This prevents assigning the same target twice.
    ' --------------------------------------------------------
    Dim tempVal As String

    For i = dayFirstRow To dayLastRow
        If Val(ws.Cells(i, COL_DAY).Value) = currentDay Then
            tempVal = Trim$(ws.Cells(i, COL_ASSIGNED_TARGET).Text)

            If Len(tempVal) > 0 Then
                If Not assignedRegistry.Exists(tempVal) Then
                    assignedRegistry.Add tempVal, True
                End If
            End If
        End If
    Next i

    ' --------------------------------------------------------
    ' Count players who still need targets and targets still free.
    ' --------------------------------------------------------
    Dim currentTargetID As String

    For i = dayFirstRow To dayLastRow
        If Val(ws.Cells(i, COL_DAY).Value) = currentDay Then

            ' Player exists and has no assigned target.
            If Len(Trim$(ws.Cells(i, COL_PLAYER_NAME).Text)) > 0 Then
                If Len(Trim$(ws.Cells(i, COL_ASSIGNED_TARGET).Text)) = 0 Then
                    playerCapacity = playerCapacity + 1
                End If
            End If

            ' Target exists and is not already assigned.
            If Len(Trim$(ws.Cells(i, COL_TARGET_POWER).Text)) > 0 Then
                currentTargetID = Trim$(ws.Cells(i, COL_TARGET_ID).Text)

                If Len(currentTargetID) > 0 Then
                    If Not assignedRegistry.Exists(currentTargetID) Then
                        targetCapacity = targetCapacity + 1
                    End If
                End If
            End If

        End If
    Next i

    If playerCapacity = 0 Then
        MsgBox "All active players for War Day " & currentDay & " already have assignments.", vbInformation, "Nothing To Assign"
        GoTo CleanExit
    End If

    If targetCapacity = 0 Then
        MsgBox "No free targets remain for War Day " & currentDay & ".", vbExclamation, "Targets Exhausted"
        GoTo CleanExit
    End If

    ReDim Players(1 To 5, 1 To playerCapacity)
    ReDim Targets(1 To 4, 1 To targetCapacity)

    ' --------------------------------------------------------
    ' Load players and targets into memory arrays.
    '
    ' Players:
    '   1 = worksheet row
    '   2 = player power
    '   3 = assigned target ID
    '   4 = star strategy
    '   5 = assignment flag
    '
    ' Targets:
    '   1 = worksheet row
    '   2 = target ID
    '   3 = target power
    '   4 = assignment flag
    ' --------------------------------------------------------
    For i = dayFirstRow To dayLastRow
        If Val(ws.Cells(i, COL_DAY).Value) = currentDay Then

            ' Load players who need assignments.
            If Len(Trim$(ws.Cells(i, COL_PLAYER_NAME).Text)) > 0 Then
                If Len(Trim$(ws.Cells(i, COL_ASSIGNED_TARGET).Text)) = 0 Then

                    ws.Cells(i, COL_STRATEGY).Value = ""

                    pCount = pCount + 1
                    Players(1, pCount) = i
                    Players(2, pCount) = ExtractPower(ws.Cells(i, COL_PLAYER_POWER).Value)
                    Players(3, pCount) = ""
                    Players(4, pCount) = ""
                    Players(5, pCount) = 0
                End If
            End If

            ' Load available targets.
            If Len(Trim$(ws.Cells(i, COL_TARGET_POWER).Text)) > 0 Then
                currentTargetID = Trim$(ws.Cells(i, COL_TARGET_ID).Text)

                If Len(currentTargetID) > 0 Then
                    If Not assignedRegistry.Exists(currentTargetID) Then
                        tCount = tCount + 1
                        Targets(1, tCount) = i
                        Targets(2, tCount) = currentTargetID
                        Targets(3, tCount) = ExtractPower(ws.Cells(i, COL_TARGET_POWER).Value)
                        Targets(4, tCount) = 0
                    End If
                End If
            End If

        End If
    Next i

    If pCount = 0 Then
        MsgBox "No players were loaded for War Day " & currentDay & ".", vbInformation, "Nothing To Assign"
        GoTo CleanExit
    End If

    If tCount = 0 Then
        MsgBox "No targets were loaded for War Day " & currentDay & ".", vbExclamation, "Targets Exhausted"
        GoTo CleanExit
    End If

    ' --------------------------------------------------------
    ' Phase 1: Prefer 3-star assignments.
    ' Strongest players pick the hardest target they can handle.
    ' --------------------------------------------------------
    Sort2DArray Players, 2, False
    Sort2DArray Targets, 3, False

    For i = 1 To pCount
        If Players(5, i) = 0 Then
            For j = 1 To tCount
                If Targets(4, j) = 0 Then
                    If Targets(3, j) <= Players(2, i) + riskMargin Then

                        Players(5, i) = 1
                        Targets(4, j) = 1
                        Players(3, i) = Targets(2, j)
                        Players(4, i) = StarText(3)

                        Exit For
                    End If
                End If
            Next j
        End If
    Next i

    ' --------------------------------------------------------
    ' Phase 2: Assign remaining targets by rank.
    ' This spreads harder remaining targets across remaining players.
    ' --------------------------------------------------------
    Dim tIndex As Long
    tIndex = 1

    For i = 1 To pCount
        If Players(5, i) = 0 Then
            For j = tIndex To tCount
                If Targets(4, j) = 0 Then

                    Players(5, i) = 1
                    Targets(4, j) = 1
                    Players(3, i) = Targets(2, j)
                    Players(4, i) = StarText(2)

                    tIndex = j + 1
                    Exit For
                End If
            Next j
        End If
    Next i

    ' --------------------------------------------------------
    ' Phase 3: Improve assignments by swapping targets.
    ' The goal is to reduce large power gaps without losing stars.
    ' --------------------------------------------------------
    OptimizeAssignmentsBySwaps Players, pCount, Targets, tCount, riskMargin

    ' --------------------------------------------------------
    ' Phase 4: Normalize ranks inside each star group.
    ' Stronger players should face stronger targets within the
    ' same expected-star category.
    ' --------------------------------------------------------
    RankMatchStarGroup Players, pCount, Targets, tCount, 3
    RankMatchStarGroup Players, pCount, Targets, tCount, 2

    ' --------------------------------------------------------
    ' Recalculate final strategy and write results to the sheet.
    ' --------------------------------------------------------
    Dim assignmentsMade As Long
    Dim finalTargetPower As Double

    assignmentsMade = 0

    For i = 1 To pCount
        If Players(5, i) = 1 Then

            finalTargetPower = GetTargetPowerByID(Targets, tCount, CStr(Players(3, i)))
            Players(4, i) = GetWarStarStrategy(CDbl(Players(2, i)), finalTargetPower, riskMargin)

            ws.Cells(Players(1, i), COL_ASSIGNED_TARGET).Value = Players(3, i)
            ws.Cells(Players(1, i), COL_STRATEGY).Value = Players(4, i)

            assignmentsMade = assignmentsMade + 1
        End If
    Next i

    MsgBox assignmentsMade & " assignments generated for War Day " & currentDay & ".", vbInformation, "Assignment Complete"

CleanExit:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

ErrorHandler:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating

    MsgBox "Assignment failed: " & Err.Description, vbCritical, "Error"

End Sub

' ============================================================
' Swap optimizer
' ============================================================

Private Sub OptimizeAssignmentsBySwaps( _
    ByRef Players As Variant, _
    ByVal pCount As Long, _
    ByRef Targets As Variant, _
    ByVal tCount As Long, _
    ByVal riskMargin As Double)

    Dim swapMade As Boolean
    Dim eps As Double

    Dim a As Long
    Dim b As Long

    Dim tID1 As String
    Dim tID2 As String

    Dim tPow1 As Double
    Dim tPow2 As Double
    Dim pPow1 As Double
    Dim pPow2 As Double

    Dim oldGap1 As Double
    Dim oldGap2 As Double
    Dim newGap1 As Double
    Dim newGap2 As Double

    Dim oldStar1 As Long
    Dim oldStar2 As Long
    Dim newStar1 As Long
    Dim newStar2 As Long

    Dim oldStars As Long
    Dim newStars As Long

    Dim oldSq As Double
    Dim newSq As Double

    Dim tmpTargetID As String

    Dim maxAllowedLoss As Double
    Dim maxAcceptableGap As Double
    Dim protectStrongestThreeStar As Boolean
    Dim blockSwap As Boolean
    Dim topPlayerPower As Double

    eps = 0.000001
    maxAllowedLoss = 1.5
    maxAcceptableGap = 2.5
    protectStrongestThreeStar = True

    If pCount > 0 Then
        topPlayerPower = CDbl(Players(2, 1))
    Else
        topPlayerPower = 0
    End If

    Do
        swapMade = False

        For a = 1 To pCount - 1
            If Players(5, a) = 1 And Len(Trim$(CStr(Players(3, a)))) > 0 Then

                For b = a + 1 To pCount
                    If Players(5, b) = 1 And Len(Trim$(CStr(Players(3, b)))) > 0 Then

                        tID1 = CStr(Players(3, a))
                        tID2 = CStr(Players(3, b))

                        If StrComp(tID1, tID2, vbTextCompare) <> 0 Then

                            pPow1 = CDbl(Players(2, a))
                            pPow2 = CDbl(Players(2, b))

                            tPow1 = GetTargetPowerByID(Targets, tCount, tID1)
                            tPow2 = GetTargetPowerByID(Targets, tCount, tID2)

                            If tPow1 >= 0 And tPow2 >= 0 Then

                                oldGap1 = Abs(pPow1 - tPow1)
                                oldGap2 = Abs(pPow2 - tPow2)

                                newGap1 = Abs(pPow1 - tPow2)
                                newGap2 = Abs(pPow2 - tPow1)

                                oldStar1 = GetStarCount(CStr(Players(4, a)))
                                oldStar2 = GetStarCount(CStr(Players(4, b)))
                                oldStars = oldStar1 + oldStar2

                                If tPow2 <= pPow1 + riskMargin Then
                                    newStar1 = 3
                                Else
                                    newStar1 = 2
                                End If

                                If tPow1 <= pPow2 + riskMargin Then
                                    newStar2 = 3
                                Else
                                    newStar2 = 2
                                End If

                                newStars = newStar1 + newStar2

                                oldSq = oldGap1 * oldGap1 + oldGap2 * oldGap2
                                newSq = newGap1 * newGap1 + newGap2 * newGap2

                                blockSwap = False

                                ' Optional rule:
                                ' Do not let the strongest player lose a 3-star assignment.
                                If protectStrongestThreeStar Then
                                    If Abs(pPow1 - topPlayerPower) < eps Then
                                        If oldStar1 = 3 And newStar1 < 3 Then blockSwap = True
                                    End If

                                    If Abs(pPow2 - topPlayerPower) < eps Then
                                        If oldStar2 = 3 And newStar2 < 3 Then blockSwap = True
                                    End If
                                End If

                                ' Accept the swap only if:
                                ' - total expected stars do not decrease
                                ' - no new gap is too large
                                ' - no player gets much worse
                                ' - the total squared gap improves
                                If Not blockSwap _
                                   And newStars >= oldStars _
                                   And newGap1 <= maxAcceptableGap _
                                   And newGap2 <= maxAcceptableGap _
                                   And newGap1 <= oldGap1 + maxAllowedLoss _
                                   And newGap2 <= oldGap2 + maxAllowedLoss _
                                   And newSq < oldSq - eps Then

                                    tmpTargetID = Players(3, a)
                                    Players(3, a) = Players(3, b)
                                    Players(3, b) = tmpTargetID

                                    Players(4, a) = GetWarStarStrategy(pPow1, tPow2, riskMargin)
                                    Players(4, b) = GetWarStarStrategy(pPow2, tPow1, riskMargin)

                                    swapMade = True
                                End If

                            End If
                        End If
                    End If
                Next b
            End If
        Next a

    Loop While swapMade

End Sub

' ============================================================
' Helper functions
' ============================================================

Private Function ExtractPower(ByVal cellVal As Variant) As Double

    If IsEmpty(cellVal) Then
        ExtractPower = 0
        Exit Function
    End If

    Dim s As String
    Dim cleanStr As String
    Dim i As Long
    Dim ch As String

    s = CStr(cellVal)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If ch Like "[0-9]" Or ch = "." Or ch = "," Then
            cleanStr = cleanStr & ch
        End If
    Next i

    If Len(cleanStr) = 0 Then
        ExtractPower = 0
    Else
        cleanStr = Replace(cleanStr, ",", ".")
        ExtractPower = Val(cleanStr)
    End If

End Function

Private Sub Sort2DArray( _
    ByRef arr As Variant, _
    ByVal sortCol As Long, _
    ByVal ascending As Boolean)

    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim temp As Variant
    Dim itemCount As Long
    Dim fieldCount As Long
    Dim shouldSwap As Boolean

    itemCount = UBound(arr, 2)
    fieldCount = UBound(arr, 1)

    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount

            shouldSwap = False

            If ascending Then
                If CDbl(arr(sortCol, i)) > CDbl(arr(sortCol, j)) Then shouldSwap = True
            Else
                If CDbl(arr(sortCol, i)) < CDbl(arr(sortCol, j)) Then shouldSwap = True
            End If

            If shouldSwap Then
                For k = 1 To fieldCount
                    temp = arr(k, i)
                    arr(k, i) = arr(k, j)
                    arr(k, j) = temp
                Next k
            End If

        Next j
    Next i

End Sub

Private Function GetTargetPowerByID( _
    ByRef Targets As Variant, _
    ByVal tCount As Long, _
    ByVal targetID As String) As Double

    Dim j As Long

    GetTargetPowerByID = -1#

    For j = 1 To tCount
        If StrComp(CStr(Targets(2, j)), targetID, vbTextCompare) = 0 Then
            GetTargetPowerByID = CDbl(Targets(3, j))
            Exit Function
        End If
    Next j

End Function

Private Function GetWarStarStrategy( _
    ByVal playerPower As Double, _
    ByVal targetPower As Double, _
    ByVal riskMargin As Double) As String

    If targetPower <= playerPower + riskMargin Then
        GetWarStarStrategy = StarText(3)
    Else
        GetWarStarStrategy = StarText(2)
    End If

End Function

Private Function StarText(ByVal starCount As Long) As String

    Dim i As Long

    For i = 1 To starCount
        StarText = StarText & ChrW(&H2B50)
    Next i

End Function

Private Function GetStarCount(ByVal strategyText As String) As Long

    Dim starChar As String
    Dim i As Long
    Dim count As Long

    starChar = ChrW(&H2B50)

    For i = 1 To Len(strategyText)
        If Mid$(strategyText, i, 1) = starChar Then
            count = count + 1
        End If
    Next i

    GetStarCount = count

End Function

Private Sub RankMatchStarGroup( _
    ByRef Players As Variant, _
    ByVal pCount As Long, _
    ByRef Targets As Variant, _
    ByVal tCount As Long, _
    ByVal starCount As Long)

    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim tmpD As Double
    Dim tmpS As String
    Dim tmpL As Long

    ' Count players in this star group.
    n = 0

    For i = 1 To pCount
        If Players(5, i) = 1 Then
            If GetStarCount(CStr(Players(4, i))) = starCount Then
                n = n + 1
            End If
        End If
    Next i

    If n < 2 Then Exit Sub

    Dim gIdx() As Long
    Dim gPow() As Double
    Dim tID() As String
    Dim tPow() As Double

    ReDim gIdx(1 To n)
    ReDim gPow(1 To n)
    ReDim tID(1 To n)
    ReDim tPow(1 To n)

    k = 0

    For i = 1 To pCount
        If Players(5, i) = 1 Then
            If GetStarCount(CStr(Players(4, i))) = starCount Then
                k = k + 1

                gIdx(k) = i
                gPow(k) = CDbl(Players(2, i))
                tID(k) = CStr(Players(3, i))
                tPow(k) = GetTargetPowerByID(Targets, tCount, tID(k))
            End If
        End If
    Next i

    ' Sort players by power, strongest first.
    For i = 1 To n - 1
        For j = i + 1 To n
            If gPow(j) > gPow(i) Then
                tmpD = gPow(i)
                gPow(i) = gPow(j)
                gPow(j) = tmpD

                tmpL = gIdx(i)
                gIdx(i) = gIdx(j)
                gIdx(j) = tmpL
            End If
        Next j
    Next i

    ' Sort targets by power, strongest first.
    For i = 1 To n - 1
        For j = i + 1 To n
            If tPow(j) > tPow(i) Then
                tmpD = tPow(i)
                tPow(i) = tPow(j)
                tPow(j) = tmpD

                tmpS = tID(i)
                tID(i) = tID(j)
                tID(j) = tmpS
            End If
        Next j
    Next i

    ' Reassign targets rank-for-rank inside the group.
    For i = 1 To n
        Players(3, gIdx(i)) = tID(i)
    Next i

End Sub