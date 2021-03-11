Sub LoopFoldersInInbox()
    
    Dim ns                  As Object
    Dim objFolder           As Object
    Dim objSubfolder        As Object
    
    Set ns = GetObject("", "Outlook.Application").GetNamespace("MAPI")
    Set objFolder = ns.GetDefaultFolder(6) ' 6 is equal to olFolderInbox
    
    For Each objSubfolder In objFolder.Folders
        Debug.Print objSubfolder.name
        Debug.Print objSubfolder.Items.Count
    Next objSubfolder
    
End Sub
