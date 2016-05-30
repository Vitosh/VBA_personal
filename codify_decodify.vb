Option Explicit

Public Function codify(str_name) As String
    
    
    Dim l_counter           As Long
    Dim l_number            As Long
    
    Dim str_substring       As String
    Dim str_number          As String
    
    Dim str_ext             As String
    Dim str_new_ext         As String
    
    Dim str_char            As String
    Dim str_char_result     As String
    
    'making the time
    For l_counter = 1 To Len(str_name) - 3
        str_number = str_number & Mid(str_name, l_counter, 1)
    Next l_counter
    l_number = str_number
    
    'making the name
    For l_counter = 3 To 1 Step -1
    
        str_char = Mid(str_name, Len(str_name) - l_counter + 1, 1)
        str_char = Chr((Asc(str_char) + l_number) Mod LETTERS_NUMBER)
        str_char = Chr(Asc(str_char) + FIRST_ASCII)
        str_char_result = str_char_result & str_char
    
    Next l_counter
    
    codify = Hex(l_number) & str_char_result
    
End Function

Public Function decodify(str_name) As String
    
    Dim l_counter       As Long
    Dim str_char        As String
    Dim str_time        As String
    
    Dim l_left          As Long
    Dim str_right       As String
    
    'making the time
    
    For l_counter = 1 To Len(str_name) - 3
        str_time = str_time & Mid(str_name, l_counter, 1)
    Next l_counter
    
    l_left = Val("&H" & str_time)
    
    'making the name
    
    For l_counter = 3 To 1 Step -1
        str_char = Mid(str_name, Len(str_name) - l_counter + 1, 1)
        str_char = Chr(Asc(str_char) - FIRST_ASCII)
        str_right = str_right & Chr(mod_where(str_char, l_left))
        
    Next l_counter
    
    decodify = l_left & str_right

End Function

Public Function mod_where(str As String, l_left As Long) As Long
    
    Dim l_counter As Long
    
    For l_counter = 0 To LETTERS_NUMBER
        If ((l_left + l_counter + FIRST_ASCII) Mod LETTERS_NUMBER = Asc(str)) Then
            mod_where = l_counter + FIRST_ASCII
            Exit For
        End If
    Next l_counter

End Function

Public Sub PrintMyName()

    Debug.Print Chr(194) & Chr(200) & Chr(210) & Chr(206) & Chr(216)

End Sub

Public Function get_extension() As String

    get_extension = Replace(Time, ":", "") & Replace(Left(Environ("Username"), 4), ".", "")

End Function
