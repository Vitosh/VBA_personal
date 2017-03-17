'random generator string
'string generator
'string degenerator
'string code decode coder decoder codify decodify


Public Function str_generator(str_value As String, b_fix As Boolean) As String
    
    Dim l_counter   As Long
    Dim l_number    As Long
    Dim str_char    As String
    
    'On Error GoTo str_generator_Error
    
    If b_fix Then
        str_value = Left(str_value, Len(str_value) - 1)
        str_value = Right(str_value, Len(str_value) - 1)
    End If

    For l_counter = 1 To Len(str_value)
        str_char = Mid(str_value, l_counter, 1)
        If b_is_odd(l_counter) Then
            l_number = Asc(str_char) + IIf(b_fix, -2, 2)
        Else
            l_number = Asc(str_char) + IIf(b_fix, -6, 6)
        End If
        
        str_generator = str_generator + Chr(l_number)
    
    Next l_counter
    
    If Not b_fix Then
        str_generator = Chr(l_number) & str_generator & Chr(l_number)
    End If
    
    On Error GoTo 0
    Exit Function

str_generator_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure str_generator of Function Modul1"

End Function

Public Function b_is_odd(l_number As Long) As Boolean
    
    b_is_odd = l_number Mod 2

End Function
