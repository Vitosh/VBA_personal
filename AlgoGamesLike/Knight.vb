Option Explicit

Public r_range                  As Range
Public r_used_range             As Range
Public l_result                 As Long

Public Sub DeleteOthers()
    
    Dim r_cell  As Range
    
    For Each r_cell In r_used_range
        If r_cell.Interior.Color <> vbGreen Then r_cell.ClearContents
    Next r_cell
    
End Sub

Public Sub CalculatePriceWithItalic(r_cell As Range, l_size As Long, Optional b_once As Boolean = False)
    
    Dim r_row       As Range
    Dim r_col       As Range
    Dim my_cell     As Range

    Dim l_row       As Long
    Dim l_col       As Long
    
    l_result = 0
    
    'RIGHT
    l_row = r_cell.Row + 1
    l_col = r_cell.Column + 2
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)
    
    l_row = r_cell.Row - 1
    l_col = r_cell.Column + 2
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)
    
    'DOWN
    l_row = r_cell.Row + 2
    l_col = r_cell.Column + 1
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)
    
    l_row = r_cell.Row + 2
    l_col = r_cell.Column - 1
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)
    
    'LEFT
    l_row = r_cell.Row - 1
    l_col = r_cell.Column - 2
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)

    l_row = r_cell.Row + 1
    l_col = r_cell.Column - 2
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)
    
    'UP
    l_row = r_cell.Row - 2
    l_col = r_cell.Column - 1
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)

    l_row = r_cell.Row - 2
    l_col = r_cell.Column + 1
    Call CheckRow(l_row, l_col, l_size, r_cell, b_once)
    
    r_cell = l_result
    Set my_cell = Nothing

End Sub

Public Sub CheckRow(l_row As Long, l_col As Long, l_size As Long, r_cell As Range, b_once As Boolean)

    If l_row <= l_size And l_col <= l_size And l_row > 0 And l_col > 0 Then
        If Len(Cells(l_row, l_col)) < 1 And Cells(l_row, l_col).Address <> r_cell.Address Then
            l_result = l_result + 1
            If b_once Then Call CalculatePriceWithItalic(Cells(l_row, l_col), l_size)
        End If
    End If

End Sub

Sub main()

    Dim my_array()          As Variant
    Dim my_array_b()        As Variant
    
    Dim l_counter           As Long
    Dim l_counter_2         As Long
    Dim l_counter_moves     As Long: l_counter_moves = 1
    Dim my_cell             As Range
    Dim b_animate           As Boolean
    Dim l_starting_row      As Long
    Dim l_starting_col      As Long
    
    b_animate = True
    l_counter = 8
    l_starting_row = 8
    l_starting_col = 8
    
    If l_starting_row > l_counter Or l_starting_row < 1 Then l_starting_row = l_counter
    If l_starting_col > l_counter Or l_starting_col < 1 Then l_starting_col = l_counter
    
    Call OnStart(b_animate)
    
    ReDim my_array(l_counter)
    
    Set r_used_range = Range(Cells(1, 1), Cells(100, 100))
    r_used_range.Clear
    
    Set r_used_range = Range(Cells(1, 1), Cells(l_counter, l_counter))
    r_used_range.Clear
    
    
    Call FormatRangeInitially(r_used_range)
    
    For l_counter_2 = 1 To l_counter
        ReDim my_array_b(l_counter)
        my_array(l_counter_2) = my_array_b
    Next l_counter_2
    
    Set my_cell = Cells(l_starting_row, l_starting_col)
    
    While l_counter_moves <= (l_counter ^ 2)
        Call CalculatePriceWithItalic(my_cell, l_counter, True)
        Call FormatMyCell(my_cell, l_counter_moves, 1)
        
        If b_animate Then Application.Wait (Now + TimeValue("00:00:01"))
                
        Call FormatMyCell(my_cell, l_counter_moves, 2)
        
        l_counter_moves = l_counter_moves + 1
        Set my_cell = FindNextTarget
        
        Call DeleteOthers
    Wend
    
    Set r_used_range = Nothing
    Set r_range = Nothing
    Set my_cell = Nothing
    
    Call OnEnd
    
End Sub

Function FindNextTarget() As Range
    
    Dim my_next     As Range
    Dim lowest      As Long: lowest = 9999
    
    For Each my_next In r_used_range
        If my_next.Value < lowest And my_next.Value > 0 And my_next.Interior.Color <> vbGreen Then
            lowest = my_next.Value
            Set FindNextTarget = my_next
        End If
    Next my_next
    
End Function

Sub FormatMyCell(ByRef my_cell_range As Range, l_counter As Long, l_color As Long)
    
    If l_color = 2 Then my_cell_range.Interior.Color = vbGreen
    If l_color = 1 Then my_cell_range.Interior.Color = vbRed
    
    my_cell_range = l_counter

End Sub

Public Sub FormatRangeInitially(r_range As Range)
    
    r_range.HorizontalAlignment = xlCenter
    r_range.Borders(xlDiagonalDown).LineStyle = xlNone
    r_range.Borders(xlDiagonalUp).LineStyle = xlNone
    With r_range.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r_range.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r_range.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r_range.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r_range.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r_range.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    r_range.ColumnWidth = 3.2

End Sub

Public Sub OnStart(b_animate As Boolean)
    
    Application.DisplayAlerts = False
    If Not b_animate Then Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Application.EnableEvents = False

End Sub

Public Sub OnEnd()
    
    'Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
End Sub


