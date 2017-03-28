Private Sub cmdClearSchedule_Click()
    Worksheets("Schedule").Activate
    
    Range("F4:L41").Select
    Selection.ClearContents
    Range("F4").Select

End Sub
Sub XFill(what As String)
'Fills the empty schedule blocks with the string.

    Dim cell As Range
    
    Worksheets("Schedule").Activate
    Range("F4:L41").Select
    
    For Each cell In Selection
        If IsEmpty(cell) Then
            cell.Value = what
        End If
    Next
    
    Range("A1").Select
    
End Sub
Private Sub cmdColorize_Click()
'Colors the schedule

    Dim cell As Range
    
    Worksheets("Schedule").Activate
    Range("F4:L41").Select
    
    For Each cell In Selection
        If cell.Value = "3:55 DN B" Then
            cell.Font.ColorIndex = 13
        ElseIf cell.Value = "3:55 UP B" Then
            cell.Font.ColorIndex = 13
        ElseIf cell.Value = "2:55 UP B" Then
            cell.Font.ColorIndex = 13
        ElseIf cell.Value = "8:55 B" Then
            cell.Font.ColorIndex = 13
        
        ElseIf cell.Value = "9:55:00 AM" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "10:55:00 AM" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "'9:55" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "'10:55" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "9:55" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "10:55" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "'9:55:00 AM" Then
            cell.Font.ColorIndex = 10
        ElseIf cell.Value = "'10:55:00 AM" Then
            cell.Font.ColorIndex = 10
        
        ElseIf cell.Value = "3:55 BK" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "3:55 MID" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "3:55 UP" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "4:25 MEZ" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "4:25 UP" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "4:25 MID" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "4:55 MID" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "4:55 MEZ" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "5:55 UP" Then
            cell.Font.ColorIndex = 1
        
        ElseIf cell.Value = "3:55 FR/CL" Then
            cell.Font.ColorIndex = 30
        ElseIf cell.Value = "4:25 FR/CL" Then
            cell.Font.ColorIndex = 30
        ElseIf cell.Value = "4:55 FR/CL" Then
            cell.Font.ColorIndex = 30
        ElseIf cell.Value = "5:55 FR/CL" Then
            cell.Font.ColorIndex = 30
        
        ElseIf cell.Value = "4:30 H" Then
            cell.Font.ColorIndex = 23
        ElseIf cell.Value = "5 H" Then
            cell.Font.ColorIndex = 23
        ElseIf cell.Value = "5:30 H" Then
            cell.Font.ColorIndex = 23
        ElseIf cell.Value = "6 H" Then
            cell.Font.ColorIndex = 23
        ElseIf cell.Value = "5 EX" Then
            cell.Font.ColorIndex = 23
        ElseIf cell.Value = "5 FR" Then
            cell.Font.ColorIndex = 23
        ElseIf cell.Value = "6 FR" Then
            cell.Font.ColorIndex = 23
            
        ElseIf cell.Value = "X" Then
            cell.Font.ColorIndex = 1
        ElseIf cell.Value = "OFF" Then
            cell.Font.ColorIndex = 1

        End If
    Next
    
    Range("A1").Select
End Sub
Private Sub cmdBringAvailability_Click()
Set wb = ActiveWorkbook
wb.Sheets("Schedule").Range("F4:L41").Value = wb.Sheets("Availability").Range("F4:L41").Value
End Sub
Sub DupeCheck(WhatRow As String)
Dim rng As Range
Dim cel As Range
Dim DupesFound As Boolean
Dim DupeInstances As Integer

DupeInstances = 0
DupesFound = False

Set rng = Range(Range(WhatRow & "1"), Range(WhatRow & Rows.Count).End(xlUp))

For Each cel In rng
    If WorksheetFunction.CountIf(rng, cel.Value) > 1 Then
        If cel.Text = "X" Then
        ElseIf cel.Text = "OFF" Then
        ElseIf cel.Text = "RO" Then
        Else
            'cel.Interior.ColorIndex = 3
            cel.Font.ColorIndex = 3
            DupesFound = True
            DupeInstances = DupeInstances + 1
        End If
    End If
Next cel

If DupesFound = True Then
    MsgBox "Duplicates found and highlighted.  Row: " & WhatRow & " - Instances: " & DupeInstances
End If

End Sub

Private Sub cmdCheckDupes_Click()
DupeCheck "F"
DupeCheck "G"
DupeCheck "H"
DupeCheck "I"
DupeCheck "J"
DupeCheck "K"
DupeCheck "L"
End Sub

Private Sub cmdSched_Click()

Worksheets("Schedule").Select

End Sub

Private Sub cmdClear_Click()
    ClearFields
End Sub
Sub ClearFields()
    lstDead.Clear
    lstPool.Clear
    lstShifts.Clear
    lstBug.Clear
    lstCovered.Clear
End Sub

Private Sub cmdGenBlock_Click()



End Sub
Sub GenerateEmployees(wutType As String)
    If wutType = "BAR" Then
        GenerateBar
    ElseIf wutType = "SERV" Then
        GenerateServ
    ElseIf wutType = "HOST" Then
        GenerateHost
    ElseIf wutType = "FR" Then
        GenerateFR
    ElseIf wutType = "EX" Then
        GenerateEX
    ElseIf wutType = "MGR" Then
        GenerateMGR
    End If
    
End Sub
Sub XCheckShifts()
Dim lrow As Long


End Sub
Sub GenerateBlock(day As String, ShfType As String)
    Dim emp As String
    Dim shft As String
    Dim lrow As Long
    Dim AMOffset As Integer
    Dim DayOffset As Integer
    Dim loopCount As Long

    Randomize

    loopCount = 1

    GenerateEmployees ShfType
    GenShifts day, ShfType
    
    
    Worksheets("Schedule").Activate
    
    DayOffset = DoDayOffset(day)
    
    Range("A1").Select
    
    'Exit Sub
StartX:
    For x = 0 To lstShifts.ListCount
        If lstShifts.ListCount = 0 Then GoTo EndX
'StartX:
        loopCount = loopCount + 1
        If loopCount > 100 Then
            MsgBox "Loop Error"
            Exit Sub
        End If
        Range("A1").Select
        randemp = Int((lstPool.ListCount) * Rnd + 0)
        emp = lstPool.List(randemp)

        If CheckDeadPool(emp) = True Then
            lstBug.AddItem emp & " is already scheduled!"
            GoTo StartX
        End If

        randshift = Int((lstShifts.ListCount) * Rnd + 0)
        shft = lstShifts.List(randshift)
        
        If CheckCovered(shft) = True Then
            lstBug.AddItem shft & " IS COVERED!"
            lstShifts.RemoveItem randshift
            GoTo StartX
        End If
    
        If shft = "9:55" Then
            shft = "'9:55"
        ElseIf shft = "10:55" Then
            shft = "'10:55"
        End If
    
        If CheckIsAM(shft) = True Then
            AMOffset = -1
        Else
            AMOffset = 0
        End If
    
        lrow = Range("A1:A1000").Find(emp).Row
    
        Selection.Offset(lrow + AMOffset, DayOffset).Select
        If Selection.Value <> "" Then
            lstBug.AddItem "CONFLICT with " & emp & " working " & shft & " - " & lrow & " " & CheckIsAM(shft)
            GoTo StartX
        End If
        Selection.Value = shft
        lstDead.AddItem emp
        
        lstPool.RemoveItem randemp
        lstShifts.RemoveItem randshift
EndX:
    Next x
    
    If cbAutoColor.Value = True Then
        cmdColorize_Click
    End If
    
    Range("A1").Select
End Sub
Function CheckCovered(WhatShift As String) As Boolean
    Dim i As Integer

    For i = 0 To lstCovered.ListCount - 1

    If lstCovered.List(i) = WhatShift Then
        CheckCovered = True
        Exit Function
    End If

    Next i

End Function
Function CheckDeadPool(WhoDead As String) As Boolean
    Dim i As Integer

    For i = 0 To lstDead.ListCount - 1

    If lstDead.List(i) = WhoDead Then
        CheckDeadPool = True
        Exit Function
    End If

Next i
End Function
Function DoDayOffset(day As String) As Integer
    If day = "Monday" Then
        DoDayOffset = 5
    ElseIf day = "Tuesday" Then
        DoDayOffset = 6
    ElseIf day = "Wednesday" Then
        DoDayOffset = 7
    ElseIf day = "Thursday" Then
        DoDayOffset = 8
    ElseIf day = "Friday" Then
        DoDayOffset = 9
    ElseIf day = "Saturday" Then
        DoDayOffset = 10
    ElseIf day = "Sunday" Then
        DoDayOffset = 11
    End If
End Function
Function CheckIsAM(shift As String) As Boolean
'Checks if a shift is AM (returns true) or PM (returns false)

If shift = "8:55 B" Then
    CheckIsAM = True
ElseIf shift = "'9:55" Then
    CheckIsAM = True
ElseIf shift = "'10:55" Then
    CheckIsAM = True
End If

End Function
Sub GenCoveredShifts(day As String)
'This scans a particular day (or column) for days already covered.
'Generates a list of shifts listed in a column and compares it to the list of shifts
'that need to be covered for that day.  If a needed shift is already covered, it is
'removed from possible shifts, and thus not scheduled.

'THIS SHOULD NOT BE USED BEFORE SUB GenShifts IS PERFORMED

On Error GoTo ExitLoc

Worksheets("Schedule").Activate

    If day = "Monday" Then
        Range("F3").Select
    ElseIf day = "Tuesday" Then
        Range("G3").Select
    ElseIf day = "Wednesday" Then
        Range("H3").Select
    ElseIf day = "Thursday" Then
        Range("I3").Select
    ElseIf day = "Friday" Then
        Range("J3").Select
    ElseIf day = "Saturday" Then
        Range("K3").Select
    ElseIf day = "Sunday" Then
        Range("L3").Select
    End If

    For x = 1 To 60
        
        If IsShift(Selection.Value) = True Then
            lstCovered.AddItem Selection.Value
        End If
        
        Selection.Offset(1, 0).Select
    Next x

ExitLoc:
End Sub
Function IsShift(shift As String) As Boolean
'Checks if a string matches a valid shift, from the schedule block only.

    If shift = "" Then
        IsShift = False
    ElseIf UCase(shift) = "X" Then
        IsShift = False
    ElseIf shift = "OFF" Then
        IsShift = False
    ElseIf shift = "RO" Then
        IsShift = False
    Else
        IsShift = True
    End If
    
End Function
Private Sub cmdGenBar_Click()
GenerateBar
End Sub
Sub GenerateBar()
    lstPool.Clear
    
    Worksheets("Employee DB").Select
    Range("B2").Select
    
    For x = 0 To 50
        If Selection.Value = "1" Then
            lstPool.AddItem (Selection.Offset(0, -1).Value)
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub
Sub GenerateServ()
    lstPool.Clear
    
    Worksheets("Employee DB").Select
    Range("C2").Select
    
    For x = 0 To 50
        If Selection.Value = "1" Then
            lstPool.AddItem (Selection.Offset(0, -2).Value)
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub
Sub GenerateHost()
    lstPool.Clear
    
    Worksheets("Employee DB").Select
    Range("E2").Select
    
    For x = 0 To 50
        If Selection.Value = "1" Then
            lstPool.AddItem (Selection.Offset(0, -4).Value)
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub
Sub GenerateFR()
    lstPool.Clear
    
    Worksheets("Employee DB").Select
    Range("F2").Select
    
    For x = 0 To 50
        If Selection.Value = "1" Then
            lstPool.AddItem (Selection.Offset(0, -5).Value)
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub
Sub GenerateEX()
    lstPool.Clear
    
    Worksheets("Employee DB").Select
    Range("G2").Select
    
    For x = 0 To 50
        If Selection.Value = "1" Then
            lstPool.AddItem (Selection.Offset(0, -6).Value)
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub
Sub GenerateMGR()
    lstPool.Clear
    
    Worksheets("Employee DB").Select
    Range("H2").Select
    
    For x = 0 To 50
        If Selection.Value = "1" Then
            lstPool.AddItem (Selection.Offset(0, -7).Value)
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub
Sub GenShifts(day As String, Posit As String)
    Dim TypeOffset As Integer
    
    Posit = UCase(Posit)

    lstShifts.Clear
    
    Worksheets("Shifts").Select
    Range("B2").Select
    
    If day = "Monday" Then
        Range("B2").Select
        TypeOffset = -1
    ElseIf day = "Tuesday" Then
        Range("C2").Select
        TypeOffset = -2
    ElseIf day = "Wednesday" Then
        Range("D2").Select
        TypeOffset = -3
    ElseIf day = "Thursday" Then
        Range("E2").Select
        TypeOffset = -4
    ElseIf day = "Friday" Then
        Range("F2").Select
        TypeOffset = -5
    ElseIf day = "Saturday" Then
        Range("G2").Select
        TypeOffset = -6
    ElseIf day = "Sunday" Then
        Range("H2").Select
        TypeOffset = -7
    End If
    
'    For x = 0 To 50
'        If Selection.Value <> "" Then
'            lstShifts.AddItem Selection.Offset(0, TypeOffset).Value & "^" & Selection.Value
'        End If
'        Selection.Offset(1, 0).Select
'    Next x
    
    For x = 0 To 50
        If Selection.Value <> "" Then
            If Selection.Offset(0, TypeOffset).Value = Posit Then
                lstShifts.AddItem Selection.Value
            End If
        End If
        Selection.Offset(1, 0).Select
    Next x
    
    Range("B2").Select
End Sub



Private Sub cmdGenEX_Click()
GenerateEX
End Sub
Private Sub cmdGenFR_Click()
GenerateFR
End Sub

Private Sub cmdGenFri_Click()
    ClearFields
    GenCoveredShifts "Friday"
    GenerateBlock "Friday", "BAR"
    GenerateBlock "Friday", "SERV"
    GenerateBlock "Friday", "HOST"
    GenerateBlock "Friday", "FR"
End Sub

Private Sub cmdGenHost_Click()
GenerateHost
End Sub
Private Sub cmdGenMGR_Click()
GenerateMGR
End Sub

Private Sub cmdGenSat_Click()
    ClearFields
    GenCoveredShifts "Saturday"
    GenerateBlock "Saturday", "BAR"
    GenerateBlock "Saturday", "SERV"
    GenerateBlock "Saturday", "HOST"
    GenerateBlock "Saturday", "FR"
End Sub

Private Sub cmdGenServ_Click()
GenerateServ
End Sub
Private Sub cmdGenMon_Click()
    ClearFields
    GenCoveredShifts "Monday"
    GenerateBlock "Monday", "BAR"
    GenerateBlock "Monday", "SERV"
    GenerateBlock "Monday", "HOST"
End Sub

Private Sub cmdGenSun_Click()
    ClearFields
    GenCoveredShifts "Sunday"
    GenerateBlock "Sunday", "BAR"
    GenerateBlock "Sunday", "SERV"
End Sub

Private Sub cmdGenThurs_Click()
    ClearFields
    GenCoveredShifts "Thursday"
    GenerateBlock "Thursday", "BAR"
    GenerateBlock "Thursday", "SERV"
End Sub

Private Sub cmdGenTue_Click()
    ClearFields
    GenCoveredShifts "Tuesday"
    GenerateBlock "Tuesday", "BAR"
    GenerateBlock "Tuesday", "SERV"
    GenerateBlock "Tuesday", "HOST"
End Sub

Private Sub cmdGenWed_Click()
    ClearFields
    GenCoveredShifts "Wednesday"
    GenerateBlock "Wednesday", "BAR"
    GenerateBlock "Wednesday", "SERV"
    GenerateBlock "Wednesday", "HOST"
    GenerateBlock "Wednesday", "FR"
    GenerateBlock "Wednesday", "EXPO"
End Sub


Private Sub cmdXFill_Click()
XFill ("OFF")
End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub cmdOpts_Click()
If Me.Height = 340 Then
    Me.Height = 200
    Me.Width = 170
Else
    Me.Height = 340
    Me.Width = 390
End If
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Height = 200
    Me.Width = 170
    cbAutoColor.Value = True
End Sub
