'github.com/timhall/Excel-TDD
'SpecExpectation
Private Function IsEqual(Actual As Variant, Expected As Variant) As Variant
    
    Dim l_count     As Long
    
    'here vitosh
    If IsArray(Expected) Then
        For l_count = LBound(Expected) To UBound(Expected)
            If Not Expected(l_count) = Actual(l_count) Then
                Debug.Print l_count
                IsEqual = False
                Exit Function
            End If
        Next l_count
    End If
    'end 
    
    If IsError(Actual) Or IsError(Expected) Then
        IsEqual = False
    ElseIf IsObject(Actual) Or IsObject(Expected) Then
        IsEqual = "Unsupported: Can't compare objects"
    ElseIf VarType(Actual) = vbDouble And VarType(Expected) = vbDouble Then
        ' It is inherently difficult/almost impossible to check equality of Double
        ' http://support.microsoft.com/kb/78113
        '
        ' Compare up to 15 significant figures
        ' -> Format as scientific notation with 15 significant figures and then compare strings
        IsEqual = IsCloseTo(Actual, Expected, 15)
    Else
        IsEqual = Actual = Expected
    End If
End Function
