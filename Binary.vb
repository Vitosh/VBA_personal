Function LngToBinary(ByVal n As Long) As String

    Dim k As Long

    LngToBinary = vbNullString
    
    If n < -2 ^ 15 Then
        LngToBinary = "0"
        n = n + 2 ^ 16
        k = 2 ^ 14
        
    ElseIf n < 0 Then
        
        LngToBinary = "1"
        n = n + 2 ^ 15
        k = 2 ^ 14
    
    Else
        
        k = 2 ^ 15
    
    End If

    Do While k >= 1
        LngToBinary = LngToBinary & Fix(n / k)
        n = n - k * Fix(n / k)
        k = k / 2
    Loop
    
End Function
