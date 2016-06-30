Option Explicit

Sub myPathForFolder()
    Debug.Print GetFolder(Environ("USERPROFILE"))
End Sub

Function GetFolder(InitialLocation As String) As String

    On Error GoTo GetFolder_Error

    Dim FolderDialog        As FileDialog
    Dim SelectedFolder      As String
    
    Set FolderDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)
    
    With FolderDialog
        .Title = "My Title For Dialog"
        .AllowMultiSelect = False
        .InitialFileName = InitialLocation
        If .Show <> -1 Then GoTo GetFolder_Error
        SelectedFolder = .SelectedItems(1)
    End With
    
    GetFolder = SelectedFolder
    Set FolderDialog = Nothing
    
    On Error GoTo 0
    Exit Function

GetFolder_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure GetFolder of Function Modul1"
    Set FolderDialog = Nothing
    
End Function
