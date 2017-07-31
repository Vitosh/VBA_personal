'---------------------------------------------------------------------------------------
' Method : MakeAllValues
' Author : v.doynov
' Date   : 07.11.2016
' Purpose: Select the range, for which you want the TDD code.
' Make sure that you can compile!!! (CreateLogFile and change_commas)
'---------------------------------------------------------------------------------------
Public Sub MakeAllValues()

    Dim my_cell                 As Range
    Dim l_counter               As Long
    Dim str                     As String
    Dim str_result              As String
    
    STR_ERROR_REPORT = ""
    
    For Each my_cell In Selection
        Call Increment(l_counter)
        str = vbTab & "my_arr(" & l_counter & ")= "
        
        If Len(my_cell) > 0 Then
            If IsDate(my_cell) Then
                str = str & "CDate(""" & my_cell & """)"
            Else
                If Not IsNumeric(my_cell) Then
                    str = str & """" & my_cell & """"
                Else
                    str = str & change_commas(my_cell.value)
                End If
            End If
        Else
            If my_cell.HasFormula Then
                str = str & """"""
            Else
                str = str & 0
            End If
        End If
        
        If Len(str_result) = 0 Then
            str_result = str
        Else
            str_result = str_result & vbCrLf & str
        End If
        
    Next my_cell
    
    Debug.Print str_result
    Call CreateLogFile(str_result)

End Sub
