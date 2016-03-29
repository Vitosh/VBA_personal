Option Explicit

Sub RemoveFormulasFromAnotherSheet()
    
    Dim rng_cell            As Range
    Dim str_inside          As String: str_inside = ":\"
    
    For Each rng_cell In ActiveSheet.UsedRange 'Selection
        If InStr(rng_cell.Formula, str_inside) > 0 Then
            Debug.Print rng_cell.Formula
            Debug.Print rng_cell.Address
            Debug.Print "---------------------------"
            'rng_cell.Value = rng_cell.Value
        End If
    Next rng_cell
End Sub

Sub ExtendContentFromRight()
    
    Dim rng_first           As Range

    Set rng_first = Selection.Cells(1, 1)
    
    Selection.Formula = rng_first.Formula
    
    Set rng_first = Nothing
    
 End Sub

Public Sub ColorSS()
    
    On Error GoTo ColorSS_Error
    
    'Colors Saturdays and Sundays.
    
    Dim r_cell      As Range
    Dim r_range     As Range
    
    For Each r_cell In Selection
        If Weekday(r_cell.Value) = 1 Or Weekday(r_cell.Value) = 7 Then
            Set r_range = ActiveSheet.Range(Cells(4, r_cell.Column), Cells(667, r_cell.Column))
            r_range.Interior.Color = 13434828
        End If
    Next r_cell
    
    Set r_range = Nothing

    On Error GoTo 0
    Exit Sub

ColorSS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ColorSS of Sub mod_play_with_me"
End Sub

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
