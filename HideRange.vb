Public Sub HideRange(r_range_to_hide As Range, l_ba_value As Long)

    Dim my_cell             As Range

    For Each my_cell In r_range_to_hide
        If my_cell.Row > l_ba_value Then
            my_cell.Interior.Pattern = xlGray8
            
            my_cell.Font.ThemeColor = xlThemeColorDark1
            
        Else
            my_cell.Interior.Pattern = xlAutomatic
            my_cell.Font.ColorIndex = xlAutomatic
        End If
    Next my_cell
     
    r_range_to_hide.Borders(xlEdgeTop).LineStyle = xlContinuous
    r_range_to_hide.Borders(xlEdgeLeft).LineStyle = xlContinuous
    r_range_to_hide.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r_range_to_hide.Borders(xlEdgeRight).LineStyle = xlContinuous

End Sub
