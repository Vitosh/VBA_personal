Option Explicit
Implements IGeneral

Private Sub IGeneral_Info()
    Debug.Print "The TG are deep!"
End Sub

Private Function IGeneral_CalculatePrice(ByVal dbl_price As Double) As Double
    IGeneral_CalculatePrice = dbl_price * -1
End Function
