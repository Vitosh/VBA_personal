Private Sub tbx_hex_Change()
    
    On Error Resume Next
    
    Dim s_write     As String
    Dim s_hour$, s_min$, s_sec$
    
    Me.lbl_hex = Val("&H" & Me.tbx_hex)
    
    
    If Len(Me.lbl_hex) = 6 Then
        s_hour = Left(Me.lbl_hex, 2)
        s_min = Mid(Me.lbl_hex, 3, 2)
        Debug.Print Me.lbl_hex
        Debug.Print s_min
        
        s_sec = Right(Me.lbl_hex, 2)
        
        s_write = s_hour & ":" & s_min & ":" & s_sec
        Me.lbl_hex = s_write
        
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub UserForm_Activate()

    Dim l_files As Long

    With Me
        .Top = CLng((Application.Height / 2 + Application.Top) - .Height / 2)
        .Left = CLng((Application.Width / 2 + Application.Left) - .Width / 2)
    End With
    
    Me.BackColor = ActiveSheet.Tab.Color
    
    If (ActiveSheet.Tab.Color = False) Then Unload Me
        
    frm_run.tbx_hex.Visible = b_value_in_array(Environ("Username"), ADMINS, True)
    frm_run.lbl_hex.BackColor = ActiveSheet.Tab.Color
    
    l_files = lng_files_to_create
    
    If l_files = 1 Then
        frm_run.lbl_hex = l_files & " Datei zu generieren."
    Else
        frm_run.lbl_hex = l_files & " Dateien zu generieren."
    End If
    
End Sub
