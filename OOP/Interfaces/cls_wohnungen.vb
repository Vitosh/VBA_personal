Option Explicit
Implements IUnitTypes

Public Sub IUnitTypes_Info()

    Debug.Print "Price is " & 2000
    
End Sub

Public Sub IUnitTypes_WriteTypes()

    Debug.Print STR_VS
    Debug.Print STR_GF
    Debug.Print STR_SF
    
End Sub

Public Sub IUnitTypes_WriteOn(str_name As String)

    Debug.Print "I am wohnung and as you wish, " & str_name
    
End Sub

Public Function IUnitTypes_CalculatePrice(dbl_m2 As Double, dbl_price_per_m2 As Double) As Double
    
    IUnitTypes_CalculatePrice = dbl_m2 * dbl_price_per_m2 + 1000
    
End Function
