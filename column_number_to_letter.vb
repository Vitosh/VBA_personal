Public Function convert_number_to_letter(ByVal l_number As Long) As String

    Dim iAlpha As Integer
    Dim iRemainder As Integer
    
    iAlpha = Int(l_number / 27)
    iRemainder = l_number - (iAlpha * 26)
    If iAlpha > 0 Then
        convert_number_to_letter = Chr(iAlpha + 64)
    End If
    If iRemainder > 0 Then
        convert_number_to_letter = convert_number_to_letter & Chr(iRemainder + 64)
    End If

End Function
