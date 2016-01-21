Private Sub chb_name_Click()
    
    txtbox_name.Enabled = Not txtbox_name.Enabled
    
End Sub

Private Sub cmd_browse_Click()
    
    Dim str_file As String
    
    str_file = Application.GetOpenFilename _
        (Title:="Please choose a file to open", _
        FileFilter:="Excel Files *.xls* (*.xls*),")
    
    txtbox_display.Caption = str_file
    
End Sub

Private Sub cmd_MainGen_Click()
    
    Call MainGen
    
End Sub
