Public Sub ShowErrors()
    
    Dim my_cell             As Range
    Dim str_result          As String
    
    For Each my_cell In ActiveSheet.UsedRange
        If IsError(my_cell) Then
            str_result = str_result & vbCrLf & my_cell.Address
        End If
    Next my_cell
    
    If Len(str_result) > 1 Then MsgBox str_result
    
End Sub
