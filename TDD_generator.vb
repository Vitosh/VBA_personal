
'---------------------------------------------------------------------------------------
' Method : MakeAllValues
' Author : v.doynov
' Date   : 03.11.2016
' Purpose: Select the range, for which you want the TDD code.
'---------------------------------------------------------------------------------------
Public Sub MakeAllValues()
    
    Dim my_cell                 As Range
    Dim l_counter               As Long
    Dim str                     As String
    
    For Each my_cell In Selection
        Call Increment(l_counter)
        str = "my_arr(" & l_counter & ")= "
        
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
            str = str & 0
        End If
        
        Debug.Print str
    Next my_cell
    
End Sub
