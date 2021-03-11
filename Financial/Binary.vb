Option Explicit
Option Private Module

Public Sub TestMe()
        
    Dim arrProducts     As Variant
    Dim lngCounter      As Long
    Dim lngValue        As Long
    Dim strBinary       As String
    Dim lngNumber       As Long
    
    arrProducts = Array("AAA", "BBB", "CCC", "DDD", "EEE", "FFF", "GGG")
                           '1,     2,     4,     8,    16,    32,    64
    lngNumber = 65 '1+2+8+16
    strBinary = StrReverse(LngToBinary(lngNumber))
    
    For lngCounter = 1 To Len(strBinary)
        lngValue = Mid(strBinary, lngCounter, 1)
        
        If lngValue Then
            Debug.Print arrProducts(lngCounter - 1)
        End If
        
    Next lngCounter
    
End Sub

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
