'---------------------------------------------------------------------------------------
' Procedure : FixSums
' Author    : v.doynov
' Date      : 18.09.2015
' Purpose   : Fixes the formulas in the sums as per the *******.
'---------------------------------------------------------------------------------------
Public Sub FixSums(ByRef r_summen As Range, ByVal l_ba_value As Long)
    
    Dim my_cell                 As Range
    
    For Each my_cell In r_summen
        my_cell.FormulaR1C1 = "=SUM(R[-10]C:R[-" & 10 - l_ba_value + 1 & "]C)"
    Next my_cell
    
End Sub
