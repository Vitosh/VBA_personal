Option Explicit
Public my_choice As cls_arrChoice
'vitosh
Sub Load_Data_To_Object()
    
    Dim s_data As String

    Set my_choice = New cls_arrChoice
    
    
    If tbl_Input.opt_publikum Then
        my_choice.Investor = [set_abbreviation_pub]
    ElseIf tbl_Input.opt_institutionen Then
        my_choice.Investor = [set_abbreviation_insti]
    End If
    
    
    If tbl_Input.opt_de Then
        my_choice.Standort = [set_abbreviation_ger]
    ElseIf tbl_Input.opt_os Then
        my_choice.Standort = [set_abbreviation_aus]
    ElseIf tbl_Input.opt_fr Then
        my_choice.Standort = [set_abbreviation_fra]
    End If
    
    
    If tbl_Input.opt_stadt1 And tbl_Input.opt_stadt1 = [set_vie_name] Then
        my_choice.Region = [set_vie_name]
    Else
        If tbl_Input.opt_stadt1 Then
            my_choice.Region = [set_muc_name]
        ElseIf tbl_Input.opt_stadt2 Then
            my_choice.Region = [set_han_name]
        ElseIf tbl_Input.opt_stadt3 Then
            my_choice.Region = [set_bln_name]
        ElseIf tbl_Input.opt_stadt4 Then
            my_choice.Region = [set_nbg_name]
        ElseIf tbl_Input.opt_stadt5 Then
            my_choice.Region = [set_ffm_name]
        End If
    End If
    
    
    If tbl_Input.opt_wohnung Then
        my_choice.Project = [set_abbreviation_wohnungen]
    ElseIf tbl_Input.opt_gewerbe Then
        my_choice.Project = [set_abbreviation_gewerbe]
    ElseIf tbl_Input.opt_wohnung Then
        my_choice.Project = [set_abbreviation_beides]
    End If
    
    my_choice.BAnumber = tbl_Input.cb_ba_number
    my_choice.GlobalProject = tbl_Input.chb_global
    
End Sub

Sub Display_Data_From_Object()

    Debug.Print my_choice.Investor
    Debug.Print my_choice.Standort
    Debug.Print my_choice.Region
    Debug.Print my_choice.Project
    Debug.Print my_choice.BAnumber
    Debug.Print my_choice.GlobalProject
    Debug.Print my_choice.GewerbeGlobal
    
    'Set my_choice = Nothing
End Sub
