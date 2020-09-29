Option Explicit

Sub MakeSelectionWithCells(my_range As Range)

    Dim l_line_style        As Long: l_line_style = 1
    Dim l_theme_color       As Long: l_theme_color = 2
    Dim d_tint_shade        As Double: d_tint_shade = 0.349986266670736
    Dim l_weight            As Long: l_weight = 2
    Dim l_counter           As Long
    
    For l_counter = 7 To 12
        Call MakeSelectionWithCells_Separated(l_line_style, l_theme_color, d_tint_shade, l_weight, l_counter, my_range)
    Next l_counter
    
End Sub

Public Sub MakeSelectionWithCells_Separated(l_line_style As Long, _
                                            l_theme_color As Long, _
                                            d_tint_shade As Double, _
                                            l_weight As Long, _
                                            l_counter As Long, _
                                            my_range As Range)
                                            
    With my_range.Borders(l_counter)
        .LineStyle = l_line_style
        .ThemeColor = l_theme_color
        .TintAndShade = d_tint_shade
        .Weight = l_weight
    End With
    
End Sub

Public Sub BorderMe(myRange As Range)

    Dim cnt As Long

    For cnt = 7 To 10 '7 to 10 are the magic numbers for xlEdgeLeft etc
        With myRange.Borders(cnt)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next

End Sub

Public Sub FixTableWithLines(tbl As Worksheet, Optional myStep As Long = 4, Optional myStart As Long = 2)
    
    OnStart
    
    Dim i As Long
    Dim myLastRow As Long: myLastRow = LastRow(tbl.Name)
    Dim myLastColumn As Long: myLastColumn = LastColumn(tbl.Name)
    Dim myRange As Range
    
    For i = myStart + myStep To myLastRow + myStep Step myStep
        With tbl
            Set myRange = .Range(.Cells(i, 1), .Cells(i, myLastColumn))
            With myRange.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End With
    Next i
    
End Sub

