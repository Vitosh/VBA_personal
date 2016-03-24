Option Explicit

Sub ExtendContentFromRight()

    Dim rng_first           As Range
    
    Set rng_first = Selection.Cells(1, 1)
    Selection.Formula = rng_first.Formula
    Set rng_first = Nothing

End Sub

Sub TakeFontAndBackground_FromAbove(Optional l_row As Long = 1)
    
    Dim my_cell           As Range
    
    For Each my_cell In Selection
        my_cell.Font.Color = my_cell.Offset(l_row, 0).Font.Color
        my_cell.Interior.Color = my_cell.Offset(l_row, 0).Interior.Color
    Next my_cell
    
End Sub
Sub RemoveFormulasFromAnotherSheet()
    
    Dim rng_cell            As Range
    Dim str_inside          As String: str_inside = ":\"
    
    Stop 'Just to make sure
    
    For Each rng_cell In ActiveSheet.UsedRange 'Selection
        If InStr(rng_cell.Formula, str_inside) > 0 Then
            Debug.Print rng_cell.Formula
            Debug.Print rng_cell.Address
            Debug.Print "---------------------------"
            rng_cell.Value = rng_cell.Value
        End If
    Next rng_cell

End Sub

Public Sub ColorSS()

    On Error GoTo ColorSS_Error

    'Colors Saturdays and Sundays.

    Dim r_cell      As Range
    Dim r_range     As Range

    For Each r_cell In Selection
        If Weekday(r_cell.Value) = 1 Or Weekday(r_cell.Value) = 7 Then
            Set r_range = ActiveSheet.Range(Cells(4, r_cell.Column), Cells(340, r_cell.Column))
            r_range.Interior.Color = 13434828
        End If
    Next r_cell

    Set r_range = Nothing
    
    On Error GoTo 0
    Exit Sub

ColorSS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ColorSS of Sub mod_play_with_me"
End Sub

Public Sub Insert186Rows()

    Dim l_counter   As Long
    
    Stop 'Make sure it is saved and the 2.2. is selected...

    For l_counter = 1 To 186
        ActiveCell.Offset(1).EntireRow.Insert
    Next l_counter
    
End Sub

Public Sub ImmediateMe()

    ActiveSheet.Rows.Ungroup

End Sub


!!! - On a sheet!
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
   Target.EntireRow.Select
End Sub
