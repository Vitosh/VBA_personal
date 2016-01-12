Option Explicit


Public Const STR_DOT = "."
Public Const STR_DASH = "-"
Public Const STR_STAR = "*"
Public Const STR_PLUS = "+"
Public Const STR_SPACE = " "
Public Const STR_VER = "|"
Public s_final As String
'

Public Function draw_a_house(k As Long) As String
'works pretty good with Consolas or SimHei font in Excel

    Dim l_counter       As Long
    Dim l_lines         As Long
    Dim s_lines         As String
    Dim s_result        As String
    
    s_final = ""
    For l_counter = 1 To k
        s_result = string_builder(k - l_counter, STR_DOT)
        s_result = s_result + string_builder_special(l_counter * 2 - 1, k)
        s_result = s_result + string_builder(k - l_counter, STR_DOT)
        s_final = s_final + s_result & vbCrLf
    Next l_counter
    
    s_final = s_final + build_ceiling_and_floor(k) & vbCrLf
    For l_counter = 1 To k - 2
        s_final = s_final + build_in_between(k) & vbCrLf
    Next l_counter
    s_final = s_final + build_ceiling_and_floor(k)
    
    draw_a_house = s_final
        
        
        
End Function

Public Function build_in_between(k As Long) As String
    
    Dim s_result        As String
    Dim l_counter       As Long
    
    s_result = string_builder(1, STR_VER)
    For l_counter = 1 To k * 2 - 3
        s_result = s_result + string_builder(1, STR_SPACE)
    Next l_counter
    
    build_in_between = s_result + string_builder(1, STR_VER)
    
End Function

Public Function build_ceiling_and_floor(k As Long) As String
    
    Dim s_result As String
    
    s_result = string_builder(1, STR_PLUS)
    s_result = s_result + string_builder(k * 2 - 3, STR_DASH)
    
    build_ceiling_and_floor = s_result + string_builder(1, STR_PLUS)
    
End Function

Public Function string_builder_special(lines_number As Long, k As Long) As String
    
    Dim l_counter As Long
    
    If lines_number = k * 2 - 1 Then
        For l_counter = 1 To k * 2 - 1
            string_builder_special = string_builder_special + IIf(l_counter Mod 2 = 1, STR_STAR, STR_SPACE)
        Next l_counter
        Exit Function
    End If
    
    string_builder_special = string_builder(1, STR_STAR)
    If lines_number > 1 Then
        string_builder_special = string_builder_special + string_builder(lines_number - 2, STR_SPACE)
        string_builder_special = string_builder_special + string_builder(1, STR_STAR)
    End If
    
End Function

Public Function string_builder(lines_number As Long, symbol As String) As String
    
    Dim counter         As Long
    
    For counter = 1 To lines_number
        string_builder = string_builder + symbol
    Next counter
    
End Function
