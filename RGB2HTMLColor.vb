Public Function RGB2HTMLColor(R As Byte, G As Byte, _
                            B As Byte) As String


'INPUT: Numeric (Base 10) Values for R, G, and B)

'RETURNS:
'A string that can be used as an HTML Color
'(i.e., "#" + the Hexadecimal equivalent)

'For VBA the RGB is reversed. R and B are revered...

    Dim HexR, HexB, HexG As Variant
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



    RGB2HTMLColor = "#" & HexR & HexG & HexB
ErrorHandler:
End Function
