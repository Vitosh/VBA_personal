Private Sub opt_de_Click()
    'Make a userForm to keep the pictures there!
    img_flag.Picture = user_form_pics.flag_de.Picture
    opt_stadt1 = True
 
    opt_stadt2.Visible = True
    opt_stadt3.Visible = True
    opt_stadt4.Visible = True
    opt_stadt5.Visible = True
    opt_stadt1.Caption = [set_muc_name]
    opt_stadt2.Caption = [set_han_name]
    opt_stadt3.Caption = [set_bln_name]
    opt_stadt4.Caption = [set_nbg_name]
    opt_stadt5.Caption = [set_ffm_name]
    
    Call opt_stadt1_Click
    
    FixInputSheet
    
End Sub
