Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Cancel = False
    
    ThisWorkbook.Save
    Application.DisplayAlerts = False
    Call HideNeeded
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
    Application.DisplayAlerts = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    ActiveSheet.PageSetup.BlackAndWhite = False
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)

    paku_message_title = tbl_settings.Range("AJ8")
    
    If Not tbl_settings.Visible Then
        With Application
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            Sh.Delete
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
        End With
        
        MsgBox (Environ("UserName") & ", Sie können Blätter nicht hinzufügen."), vbInformation, paku_message_title
    End If
    
End Sub

Private Sub Workbook_Open()

    Call HideNeeded
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    [set_root_user] = False
    Application.Caption = ""
    
End Sub
