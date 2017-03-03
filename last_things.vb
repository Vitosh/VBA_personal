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

Public Function l_locate_value_row(target As String, ByRef target_sheet As Worksheet, _
                                   Optional l_col As Long = 1, _
                                   Optional l_more_values_found As Long = 1, _
                                   Optional b_look_for_part = False) As Long

    Dim l_values_found  As Long
    Dim r_local_range   As Range
    Dim my_cell         As Range
    
    l_values_found = l_more_values_found

    Set r_local_range = target_sheet.Range(target_sheet.Cells(1, l_col), target_sheet.Cells(Rows.Count, l_col))

    For Each my_cell In r_local_range

        'The b_look_for_part is for the vertriebscase
        If b_look_for_part Then
            If target = Left(my_cell, Len(target)) Then
                If l_values_found = 1 Then
                    l_locate_value_row = my_cell.Row
                    Exit Function
                Else
                    Call Decrement(l_values_found)
                End If
            End If
        Else
            If target = Trim(my_cell) Then
                If l_values_found = 1 Then
                    l_locate_value_row = my_cell.Row
                    Exit Function
                Else
                    Call Decrement(l_values_found)
                End If
            End If
        End If
    Next my_cell

    l_locate_value_row = -1

End Function

Public Function l_locate_value_col(target As String, _
                                    ByRef target_sheet As Worksheet, _
                                    Optional l_row As Long = 1)

    Dim cell_to_find                As Range
    Dim r_local_range               As Range
    Dim my_cell                     As Range
    
    Set r_local_range = target_sheet.Range(target_sheet.Cells(l_row, 1), target_sheet.Cells(l_row, Columns.Count))
    
    For Each my_cell In r_local_range
        If target = Trim(my_cell) Then
            l_locate_value_col = my_cell.Column
            Exit Function
        End If
    Next my_cell
    
    l_locate_value_col = -1

End Function

            
            
Public Function LastUsedColumn() As Long
    
    Dim rLastCell As Range
    
    Set rLastCell = ActiveSheet.Cells.Find(What:="*", _
                                    After:=ActiveSheet.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    
    LastUsedColumn = rLastCell.Column

End Function

Public Function LastUsedRow() As Long

    Dim rLastCell As Range

    Set rLastCell = ActiveSheet.Cells.Find(What:="*", _
                                    After:=ActiveSheet.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

    LastUsedRow = rLastCell.Row

End Function
