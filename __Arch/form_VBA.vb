Private Sub UserForm_Activate()

    img_sad.Visible = False
    img_smile.Visible = True
    
    With frm_green
        .Top = Application.Top + 200
        .Left = Application.Left + 100
    End With
    
    If b_is_error Then
    
        frm_green.lbl_status.BackColor = RGB(200, 10, 10)
        frm_green.lbl_status = [set_paku_thankyou] & vbCrLf & "Status: Nicht erfolgreich! :("
        
        img_sad.Visible = True
        img_smile.Visible = False
        Me.Repaint
        Application.Wait (Now + TimeValue("00:00:01"))
    Else
        
        frm_green.lbl_status.BackColor = RGB(10, 200, 10)
        frm_green.lbl_status = [set_paku_thankyou] & vbCrLf & "Status: Erfolgreich! :) "
    
    End If
    
    Me.Repaint
    Application.Wait (Now + TimeValue("00:00:02"))
    Unload Me
    
End Sub
