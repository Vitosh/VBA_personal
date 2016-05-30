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
    
    codify = Hex(l_number) & StrReverse(str_char_result)
    
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
    
    decodify = l_left & StrReverse(str_right)

End Function

Public Function format_decodify(str_input As String, Optional b_for_file_name As Boolean = False) As String
    
    Dim str_exchange1   As String: str_exchange1 = ":"
    Dim str_exchange2   As String: str_exchange2 = " "
    
    If b_for_file_name Then
        If Len(str_input) = 9 Then
            format_decodify = insert_in_position(str_input, str_exchange2, 6)
        Else
            format_decodify = insert_in_position(str_input, str_exchange2, 5)
        End If
        
        Exit Function
        
    End If
    
    If Len(str_input) = 9 Then
        format_decodify = insert_in_position(str_input, str_exchange1, 2)
        format_decodify = insert_in_position(format_decodify, str_exchange1, 5)
        format_decodify = insert_in_position(format_decodify, str_exchange2, 8)
    Else
        format_decodify = insert_in_position(str_input, str_exchange1, 1)
        format_decodify = insert_in_position(format_decodify, str_exchange1, 4)
        format_decodify = insert_in_position(format_decodify, str_exchange2, 7)
    End If
    
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

Public Function get_extension() As String

    get_extension = Replace(Time, ":", "") & Replace(Left(Environ("Username"), 4), ".", "")

End Function

Function insert_in_position(source As String, str As String, l As Long) As String
    'insert in position
    
    insert_in_position = Mid(source, 1, l) & str & Mid(source, l + 1, Len(source) - l)
    
End Function

Function delete_in_position(source As String, l As Long) As String
    'delete in position
    
    delete_in_position = Mid(source, 1, l - 1) & Mid(source, l + 1, Len(source) - l)
    
End Function
