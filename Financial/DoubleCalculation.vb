'double inaccuracy example example double inaccuracy floating point accuracy

Sub TestMe()
    
    Dim a           As Double: a = 20
    Dim b           As Double: b = 0.1
    
    Cells.Clear
    Range("A1") = a - b
    Range("A2") = a + b
    
    Range("A3").Formula = "=A1-A2"
    Range("A4") = b * 2 * -1
    Range("A5").Formula = "=A3=A4"
    
End Sub
