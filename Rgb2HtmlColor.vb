Option Explicit
'RGB2HTMLColor html htmlcolor
'INPUT: Numeric (Base 10) Values for R, G, and B)
'OUTPUT:
'String to be used for color of element in VBA.
'E.G -> if the color is like this:-> &H80000005&
'we should change just the last 6 positions to get our color! H80 must stay.

Public Function RGB2HTMLColor(B As Byte, G As Byte, R As Byte) As String

    Dim HexR As Variant, HexB As Variant, HexG As Variant
    Dim sTemp As String

    On Error GoTo ErrorHandler

    'R
    HexR = Hex(R)
    If Len(HexR) < 2 Then HexR = "0" & HexR

    'Get Green Hex
    HexG = Hex(G)
    If Len(HexG) < 2 Then HexG = "0" & HexG

    HexB = Hex(B)
    If Len(HexB) < 2 Then HexB = "0" & HexB

    RGB2HTMLColor = HexR & HexG & HexB
    Debug.Print "Enter RGB, without caring for the real colors, the function knows what it is doing."
    Debug.Print "IF 50D092 then &H0050D092&"

    Exit Function
    
ErrorHandler:
    Debug.Print "RGB2HTMLColor was not successful"
End Function

Sub GetHexFromInteriorCell()

    Worksheets(1).Cells(1, "A").Interior.Color = vbYellow
    Debug.Print Hex(Worksheets(1).Cells(1, "A").Interior.Color)  'FFFF
    Debug.Print Worksheets(1).Cells(1, "A").Interior.Color       '65535

    Dim hexColor As String
    hexColor = Right("000000" & Hex(Worksheets(1).Cells(1, "A").Interior.Color), 6)

    Debug.Print HexToRgb(hexColor)                               'FFFF00

End Sub

Public Function HexToRgb(hexColor As String) As String

    Dim red As String
    Dim green As String
    Dim blue As String

    red = Left(hexColor, 2)
    green = Mid(hexColor, 3, 2)
    blue = Right(hexColor, 2)

    HexToRgb = blue & green & red

End Function
