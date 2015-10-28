Option Explicit

Public Const STR_VS = "V. und S."
Public Const STR_GF = "G. und F."
Public Const STR_SF = "S. und F."
Public Const STR_G1 = "G. und W. - L."
Public Const STR_G2 = "G. und W. - G.W."

Sub test()

    Dim arr_units(1 To 4)   As IUnitTypes
    Dim l_counter           As Long
    Dim arr_prices(1 To 4)  As Double
    
    Set arr_units(1) = New cls_wohnungen
    Set arr_units(2) = New cls_gewerbe
    Set arr_units(3) = New cls_beide
    Set arr_units(4) = New cls_beide
    
    For l_counter = LBound(arr_units) To UBound(arr_units)
        Call arr_units(l_counter).Info
        Call arr_units(l_counter).WriteTypes
        Call arr_units(l_counter).WriteOn("PIV")
        arr_prices(l_counter) = arr_units(l_counter).CalculatePrice(10, 1)
    Next l_counter
    
    For l_counter = LBound(arr_prices) To UBound(arr_prices)
        Debug.Print arr_prices(l_counter)
    Next l_counter

End Sub
