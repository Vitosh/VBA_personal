Function Min(ParamArray values() As Variant) As Variant
    
    Dim minValue As Variant, Value As Variant
    minValue = values(0)
    
    For Each Value In values
        If Value < minValue Then minValue = Value
    Next
    
    Min = minValue
    
End Function

Function Max(ParamArray values() As Variant) As Variant
    
    Dim maxValue As Variant, Value As Variant
    maxValue = values(0)
    
    For Each Value In values
        If Value > minValue Then maxValue = Value
    Next
    
    Max = maxValue
    
End Function
