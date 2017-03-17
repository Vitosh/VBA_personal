Private Sub btn_save_as_Click()
        
    Dim b_saved As Boolean
    
    b_saved = Application.Dialogs(xlDialogSaveAs).Show
    If Not b_saved Then MsgBox "Die Datei wurde nicht gespeichert!", vbInformation, [ale]

End Sub
