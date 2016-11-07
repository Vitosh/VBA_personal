'Encrypt, encript,
'Decrypt, decript,
'password, check hours

Option Explicit

Public Const FIRST_ASCII = 97
Public Const LETTERS_NUMBER = 26

Public Function codify_time() As String

    If [set_in_production] Then On Error GoTo codify_Error
    
    Dim dbl_01                  As Variant
    Dim dbl_02                  As Variant
    Dim dbl_now                 As Double
    
    dbl_now = Round(Now(), 8)
    
    dbl_01 = Split(CStr(dbl_now), ",")(0)
    dbl_02 = Split(CStr(dbl_now), ",")(1)
    
    codify_time = Hex(dbl_01) & "_" & Hex(dbl_02)

   On Error GoTo 0
   Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function


Public Function codify(str_name) As String
    
    Dim l_counter           As Long
    Dim l_number            As Long
    
    Dim str_number          As String
    
    Dim str_char            As String
    Dim str_char_result     As String
    
    Dim str_first           As String
    Dim str_last            As String
    
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
    
    'now reverse first and last positions
    str_first = get_in_position(codify, 1)
    str_last = get_in_position(codify, 1, True)
    
    codify = delete_in_position(codify, 1)
    codify = delete_in_position(codify, Len(codify))
    
    codify = insert_in_position(codify, str_first, Len(codify))
    codify = insert_in_position(codify, str_last, 0)
    
    codify = LCase(codify)
    
End Function

Public Function decodify(str_name) As String
    
    Dim l_counter       As Long
    Dim str_char        As String
    Dim str_time        As String
    
    Dim l_left          As Long
    Dim str_right       As String
    
    Dim str_first       As String
    Dim str_last        As String
    
    'now reverse first and last positions
    str_first = get_in_position(str_name, 1)
    str_last = get_in_position(str_name, 1, True)
    
    str_name = delete_in_position(str_name, 1)
    str_name = delete_in_position(str_name, Len(str_name))
    
    str_name = insert_in_position(str_name, str_first, Len(str_name))
    str_name = insert_in_position(str_name, str_last, 0)
    
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

Function insert_in_position(ByVal source As String, str As String, l As Long) As String
    'insert in position
    
    insert_in_position = Mid(source, 1, l) & str & Mid(source, l + 1, Len(source) - l)
    
End Function

Function delete_in_position(ByVal source As String, l As Long) As String
    'delete in position
    
    delete_in_position = Mid(source, 1, l - 1) & Mid(source, l + 1, Len(source) - l)
    
End Function

Function get_in_position(ByVal str As String, l_position As Long, Optional b_is_last As Boolean = False) As String
    
    get_in_position = Mid(str, l_position, 1)
    
    If b_is_last Then get_in_position = Mid(str, Len(str), 1)
    
End Function



