Public Function NumberToLetter(number As Long) As String

On Error GoTo NumberToLetterError

    Dim remainder As Long

    If number < 1 Or number > 2 ^ 14 Then
        Err.Raise 999, Description:="Error on " & number
    End If

    Do While number > 0
       remainder = (number - 1) Mod 26
       NumberToLetter = Chr(65 + remainder) + NumberToLetter
       number = (number - remainder) \ 26
    Loop
    
    Exit Function
    
NumberToLetterError:
    NumberToLetter = Error
End Function

Public Sub NumberToLetterTest()

    Debug.Print NumberToLetter(1) = "A"
    Debug.Print NumberToLetter(26) = "Z"
    Debug.Print NumberToLetter(27) = "AA"
    Debug.Print NumberToLetter(100) = "CV"
    Debug.Print NumberToLetter(200) = "GR"
    Debug.Print NumberToLetter(701) = "ZY"
    Debug.Print NumberToLetter(702) = "ZZ"

    Debug.Print NumberToLetter(703) = "AAA"
    Debug.Print NumberToLetter(715) = "AAM"
    Debug.Print NumberToLetter(1379) = "BAA"
    Debug.Print NumberToLetter(2055) = "CAA"
    Debug.Print NumberToLetter(2731) = "DAA"
    Debug.Print NumberToLetter(704) = "AAB"
    Debug.Print NumberToLetter(1380) = "BAB"
    Debug.Print NumberToLetter(2056) = "CAB"
    Debug.Print NumberToLetter(2732) = "DAB"
    Debug.Print NumberToLetter(2812) = "DDD"
    Debug.Print NumberToLetter(5434) = "GZZ"
    Debug.Print NumberToLetter(8138) = "KZZ"
    Debug.Print NumberToLetter(16000) = "WQJ"
    Debug.Print NumberToLetter(16251) = "XAA"
    Debug.Print NumberToLetter(16384) = "XFD"

    Debug.Print NumberToLetter(16386) = "Error on 16386"
    Debug.Print NumberToLetter(-3) = "Error on -3"

End Sub


Public Function ConvertNumberToLetterExcel(number As Long) As String
        
    ConvertNumberToLetterExcel = Split(Cells(1, number).Address, "$")(1)

End Function
