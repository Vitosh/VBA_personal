Public Sub ShowErrors()
    
    Dim my_cell             As Range
    Dim str_result          As String
    
    For Each my_cell In ActiveSheet.UsedRange
        If IsError(my_cell) Then
            str_result = str_result & vbCrLf & my_cell.Address
        End If
    Next my_cell
    
    If Len(str_result) > 1 Then MsgBox str_result
    
End Sub


Public Function change_commas(ByVal myValue As Variant) As String

    Dim str_temp As String

    str_temp = CStr(myValue)
    change_commas = Replace(str_temp, ",", ".")

End Function

Public Sub EnableMySaves()

    Application.OnKey "%{F11}"
    Application.OnKey "^c"
    Application.OnKey "^v"
    Application.OnKey "^x"
    If Not b_value_in_array(Environ("username"), ADMINS, True) Then Application.EnableCancelKey = xlDisabled

End Sub

Public Sub DisableMySaves()

    Application.OnKey "^c", "DisabledCombination"
    Application.OnKey "^v", "DisabledCombination"
    Application.OnKey "^x", "DisabledCombination"
    Application.EnableCancelKey = xlInterrupt

End Sub
