Function last_col(Optional str_sheet As String, Optional row_to_check As Long = 1) As Long
    
    Dim shSheet  As Worksheet
    
        If str_sheet = vbNullString Then
            Set shSheet = ActiveSheet
        Else
            Set shSheet = Worksheets(str_sheet)
        End If
    
    last_col = shSheet.Cells(row_to_check, shSheet.Columns.Count).End(xlToLeft).Column

End Function


Function last_row(Optional str_sheet As String, Optional column_to_check As Long = 1) As Long
    
    Dim shSheet  As Worksheet
    
        If str_sheet = vbNullString Then
            Set shSheet = ActiveSheet
        Else
            Set shSheet = Worksheets(str_sheet)
        End If
    
    last_row = shSheet.Cells(shSheet.Rows.Count, column_to_check).End(xlUp).Row

End Function

Public Function locate_value_col(str_value As String, ByRef target_sheet As Worksheet) As Long
    
    Dim cell_to_find As Range
    
    Set cell_to_find = Range(target_sheet.Cells(1, 1), target_sheet.Cells(1, last_col(tbl_totals))).Find(str_value, LookIn:=xlValues)
    locate_value_col = cell_to_find.Column
    Set cell_to_find = Nothing
    
End Function
