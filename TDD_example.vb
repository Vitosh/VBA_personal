Option Explicit

Public b_print_info As Boolean

Sub TDD_Setting_2()
    
    If b_print_info Then Debug.Print 2
    With tbl_Input
        .opt_stadt2 = True
    End With
End Sub
