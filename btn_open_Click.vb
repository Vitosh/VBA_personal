'Opens the file open file open a file open a folder open folder
'Eliminates the file name and opens the folder.

Private Sub btn_open_Click()
     On Error GoTo btn_open_Click_Error
     
     Dim my_str As String
     
     my_str = tbl_input.lblDisplayTerminPlanner
     
     Call Shell("explorer.exe" & " " & Left(my_str, Len(my_str) - Len(Split(my_str, "\")(UBound(Split(my_str, "\"))))), vbNormalFocus)

btn_open_Click_Error:
    Debug.Print Err.Number & " (" & Err.Description & ")"
    
End Sub
