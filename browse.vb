Private Sub cmd_browse_Click()
    
    Dim str_file As String
    
    str_file = Application.GetOpenFilename _
        (Title:="Please choose a file to open", _
        FileFilter:="Excel Files *.xls* (*.xls*),")
    
    txtbox_display.Text = str_file
    
End Sub
