Public Function WriteString(ByVal n As Long) As String
    'Lucida Console or Consolas
    Dim v_Bus()     As Variant
    Dim s_char      As String
    
    Dim i           As Long
    Dim l_col       As Long
    Dim l_row       As Long
    
    n = n - 1
    v_Bus = Array("+------------------------+", _
                  "|......................|D|)", _
                  "|......................|.|", _
                  "|........................|", _
                  "|......................|.|)", _
                  "+------------------------+")
    
    For i = 0 To 33
        
        If i > n Then
            s_char = "#"
        Else
            s_char = "O"
        End If

        If i < 4 Then
            l_col = 0
        ElseIf i = 4 Then
            l_col = 1
        Else
            l_col = (i - 2) / 3
        End If

        If i <= 3 Then
            l_row = i
        Else
            l_row = (i - 4) Mod 3
        End If

        If (l_row = 2 And l_col <> 0) Then l_row = l_row + 1
        Mid(v_Bus(l_row + 1), (1 + l_col * 2) + 1, 1) = s_char
    Next i
    
    WriteString = draw_bus(v_Bus)

End Function

Public Function draw_bus(v_Bus As Variant) As String
    
    Dim i As Long
    For i = LBound(v_Bus) To UBound(v_Bus)
        draw_bus = draw_bus & v_Bus(i) & vbCrLf
    Next i
    
End Function

Public Sub TestBus()
    
    Dim l_counter As Long

    For l_counter = 1 To 34
        Debug.Print l_counter
        Debug.Print WriteString(l_counter)
    Next l_counter
End Sub
