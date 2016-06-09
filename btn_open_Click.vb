'Opens the file open file open a file open a folder open folder
'Eliminates the file name and opens the folder.

Private Sub btn_open_Click()
    On Error GoTo btn_open_Click_Error
     
    Dim my_str          As String
    Dim my_str2         As String

         
    my_str = tbl_input.lblDisplayTerminPlanner
    my_str2 = Left(my_str, Len(my_str) - Len(Split(my_str, "\")(UBound(Split(my_str, "\")))))
    Call Shell("explorer.exe" & " " & my_str2, vbNormalFocus)
  
btn_open_Click_Error:
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ")"

    
End Sub
