Option Explicit


'Application.Run "Personal.xlsb!DeleteName", "NAME_HERE"
Public Sub DeleteName(sName As String)

   On Error GoTo DeleteName_Error

    ActiveWorkbook.Names(sName).Delete
    
    Debug.Print sName & " is deleted!"
    
   On Error GoTo 0
   Exit Sub

DeleteName_Error:

    Debug.Print sName & " not present or some error"
    On Error GoTo 0
    
End Sub

Sub get_names_of_cells()
    
    Dim cell        As Range
    
    On Error Resume Next
    
    For Each cell In Selection
        cell = cell.Name.Name
    Next cell
    
    On Error GoTo 0
    
End Sub

Sub set_names_of_cells()

    Dim sample_range        As Range
    Dim cell                As Range
    
    Set sample_range = Selection
        
    For Each cell In sample_range
        If Not IsEmpty(cell) Then
            cell.Name = cell.Text
            cell.Clear
        End If
    Next cell

End Sub
