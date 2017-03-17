Public Sub FormatMyCell(ByRef my_cell As range, Optional b_as_currency As Boolean = False, _
                                                Optional b_as_date As Boolean = False, _
                                                Optional b_as_dark As Boolean = False, _
                                                Optional b_as_din As Boolean = False)
                                                
    If b_as_currency Then
        my_cell.NumberFormat = "#,##0.00 $"
    End If
    
    If b_as_date Then
        my_cell.NumberFormat = "[$-407]mmm/ yy;@"
    End If
    
    If b_as_dark Then
        my_cell.Interior.ThemeColor = xlThemeColorDark1
        my_cell.Interior.TintAndShade = -0.249946592608417
    End If
    
    If b_as_din Then
        my_cell.Font.Name = "DIN-Light"
    End If

End Sub
