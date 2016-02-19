Private Sub Workbook_BeforeClose(Cancel As Boolean)

   On Error GoTo Workbook_BeforeClose_Error

    Cancel = False
    
    ThisWorkbook.Save
    Application.DisplayAlerts = False
    Call HideNeeded
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
    Application.DisplayAlerts = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    ActiveSheet.PageSetup.BlackAndWhite = False
    Me.Save
    Application.OnKey "%{F11}"


   On Error GoTo 0
   Exit Sub

Workbook_BeforeClose_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_BeforeClose of Sub xl_paku"
    
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


   On Error GoTo Workbook_Open_Error

    Call HideNeeded
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", false)"
    Application.DisplayFormulaBar = False
    [set_root_user] = False
    If Not b_value_in_array(Environ("username"), ADMINS, True) Then Application.OnKey "%{F11}", ""
    Application.WindowState = xlMaximized


   On Error GoTo 0
   Exit Sub

Workbook_Open_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_Open of Sub xl_paku"
    Me.Save
    ThisWorkbook.Close
    
End Sub
