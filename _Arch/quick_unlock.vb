Public Function b_value_in_array(my_value As Variant, my_array As Variant, Optional b_is_string As Boolean = False) As Boolean

    Dim l_counter

    If b_is_string Then
        my_array = Split(my_array, ":")
    End If

    For l_counter = LBound(my_array) To UBound(my_array)
        my_array(l_counter) = CStr(my_array(l_counter))
    Next l_counter

    b_value_in_array = Not IsError(Application.Match(CStr(my_value), my_array, 0))
    
End Function


Public Sub aaa()
    'easy to write and easy to remember
    
    If Not b_value_in_array(Environ("Username"), ADMINS, True) Then
        Debug.Print "no"
        Exit Sub
    End If
    
    Call UnhideAll 'UnprotectAll is included
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    Debug.Print "a"
    
End Sub
