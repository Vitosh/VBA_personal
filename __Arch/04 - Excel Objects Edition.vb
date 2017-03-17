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

Sub DisplayCommentsInWS()

    Dim ws_target           As Worksheet
    Dim ws_source           As Worksheet
    Dim rng_rng             As Range
    Dim rng_cell            As Variant
    Dim i                   As Long: i = 2
    Dim b_comment_found     As Boolean
    
    Call OnStart
    
    Set ws_target = Sheets("Comments") 'I would love to have an error if it does not exist
    ws_target.Cells.Delete
    ws_target.Range("A1") = "Sheet"
    ws_target.Range("B1") = "Address"
    ws_target.Range("C1") = "Comment"
    ws_target.Range("D1") = "Cell value"
    ws_target.Range("E1") = "Author"
    
    On Error Resume Next
    
    For Each ws_source In ThisWorkbook.Worksheets
        Set rng_cell = ws_source.Cells.SpecialCells(xlCellTypeComments)
        
        If Not IsEmpty(rng_cell) Then
            For Each rng_rng In rng_cell
                b_comment_found = True
                
                ws_target.Range("A" & i) = ws_source.Name
                ws_target.Range("B" & i) = rng_rng.Address
                ws_target.Range("C" & i) = rng_rng.Comment.Text
                ws_target.Range("C" & i).WrapText = False
                ws_target.Range("D" & i) = rng_rng.Value
                ws_target.Range("E" & i) = rng_rng.Comment.Author
                i = i + 1
                Debug.Print "Working " & i
                
            Next rng_rng
        End If
    Next ws_source
    
    If Not b_comment_found Then
        Debug.Print "No Comments were found. Tab ""Comments"" is deleted"
        Application.DisplayAlerts = False
        ws_target.Delete
        Application.DisplayAlerts = True
    Else
        Debug.Print "End"
    End If
    
    ws_target.Columns.AutoFit
    
    Call OnEnd
    
    On Error GoTo 0
    
    Set rng_rng = Nothing
    Set ws_source = Nothing
    Set ws_target = Nothing
    Set rng_cell = Nothing
    
End Sub

Public Sub DeleteAllComments()

    Dim ws      As Worksheet
    Dim cmt     As Comment

    For Each ws In ThisWorkbook.Worksheets
        For Each cmt In ws.Comments
            Debug.Print "Comment deleted"
            cmt.Delete
        Next cmt
    Next ws

End Sub

Public Sub OnStart()

    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Application.EnableEvents = False

End Sub

Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
End Sub
