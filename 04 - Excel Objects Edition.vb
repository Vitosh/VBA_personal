'---------------------------------------------------------------------------------------
' Method : AddStringToFormula
' Author : v.doynov
' Date   : 29.03.2016
' Purpose: Call like this =>> call AddStringToFormula(")*set_teilung_ba1") or ba2
'---------------------------------------------------------------------------------------
Public Sub AddStringToFormula(s_added_str As String)

    Dim r_range     As Range
    Dim l_counter   As Long
    
   On Error GoTo AddStringToFormula_Error

    Debug.Print Selection.Address & " -> " & Selection.Parent.Name
    Stop 'Make sure you have only one sheet active in the current app
    
    For Each r_range In Selection.SpecialCells(xlCellTypeFormulas)
        r_range.Formula = "=(" & Right(r_range.Formula, Len(r_range.Formula) - 1) & s_added_str
        Debug.Print r_range.Address & " changed"
        l_counter = l_counter + 1
    Next r_range
    
    Debug.Print vbCrLf & "Total Changes: " & l_counter

   On Error GoTo 0
   Exit Sub

AddStringToFormula_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddStringToFormula of Module mod_play"
    
End Sub
