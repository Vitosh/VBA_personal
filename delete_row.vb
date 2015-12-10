Option Explicit
'call CheckAndDelete(Range("A1:A10"),1,"1")

Public Sub CheckAndDelete(r_range As Range, l_column As Long, Optional s_char As String = ".")

    Dim l_counter As Long
    Dim r_cell  As Range
    
    For l_counter = r_range.Cells(r_range.Count).Row To r_range.Cells(1, 1).Row Step -1
        Set r_cell = Cells(l_counter, l_column)
        If InStr(1, r_cell, s_char, vbTextCompare) Then
            Rows(l_counter).EntireRow.Delete
        End If
    Next l_counter
    
    Set r_cell = Nothing
    Set r_range = Nothing
    
End Sub
