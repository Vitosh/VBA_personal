Option Explicit

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
    Debug.Print "Red and Blue are reversed ... pay attention to the input in the input"

ErrorHandler:
    Debug.Print "RGB2HTMLColor was not successful"
End Function
