Option Explicit
Implements IGeneral

Public Sub IGeneral_Info()
    Debug.Print "The carports are cheaper than TG."
End Sub

Private Function IGeneral_CalculatePrice(ByVal dbl_price As Double) As Double
    IGeneral_CalculatePrice = dbl_price * 10
End Function
