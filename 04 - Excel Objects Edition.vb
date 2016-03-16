Sub ExtendContentFromRight()
    
    Dim rng_first           As Range

    Set rng_first = Selection.Cells(1, 1)
    
    Selection.Formula = rng_first.Formula
    
    Set rng_first = Nothing
    
 End Sub

Sub RemoveFormulasFromAnotherSheet()
    
    Dim rng_cell            As Range
    Dim str_inside          As String: str_inside = ":\"
    
    For Each rng_cell In ActiveSheet.UsedRange
        If InStr(rng_cell.Formula, str_inside) > 0 Then
            Debug.Print rng_cell.Formula
            Debug.Print rng_cell.Address
            Debug.Print "---------------------------"
        End If
    Next rng_cell
    
 End Sub
