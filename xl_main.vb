Option Explicit

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
    
    Me.Save
    Application.AskToUpdateLinks = True
    
    Call EnableMySaves
    
    On Error GoTo 0
    Exit Sub

Workbook_BeforeClose_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_BeforeClose of Sub xl_paku"
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    
    If Not tbl_settings.Visible Then
        With Application
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            Sh.Delete
            Application.DisplayAlerts = True
            Application.ScreenUpdating = True
        End With

        MsgBox (Environ("UserName") & ", Sie können Blätter nicht hinzufügen."), vbInformation, ThisWorkbook.Name
    End If

End Sub

Private Sub Workbook_Open()
    
    On Error GoTo Workbook_Open_Error
    
    Call LockMe
    Call HideNeeded
    Call LockScroll(Array(tbl_main.Name, "A1:X107"))
    Call MinimizeRibbon
    
    ActiveWindow.WindowState = xlMaximized
    Application.WindowState = xlMaximized
    
    'Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", false)"
    'ActiveWindow.DisplayHeadings = False
    Application.OnKey "^{W}", "DisabledCombination"
    Application.OnKey "^{w}", "DisabledCombination"
    Application.OnKey "^{E}", "InitializeFormTotals"
    Application.OnKey "^{e}", "InitializeFormTotals"

    Call CheckHowManyWbAreOpened
    
    tbl_main.Select
    'tbl_main.tb_Show = False
    tbl_main.chb_delete = False
    
    tbl_main.Cells(1, 1).Select
    ActiveWindow.Zoom = 74
    
    On Error GoTo 0
    Exit Sub

Workbook_Open_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_Open of Sub DieseArbeitsmappe", vbInformation, [set_planerkostenberechnung]

End Sub

Public Sub CheckHowManyWbAreOpened()
    On Error GoTo CheckHowManyWbAreOpened_Error

    If Workbooks.Count > 1 Then
        [set_more_instances] = True
        frmInfo.Show (vbModeless)
        frmInfo.lb_information = "Sie haben mehr als eine Instanz von Excel. Dies ist keine sehr gute Idee."
        frmInfo.Repaint
        Application.Wait (Now + TimeValue("00:00:05"))
        Unload frmInfo
    End If
        [set_more_instances] = False

   On Error GoTo 0
   Exit Sub

CheckHowManyWbAreOpened_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckHowManyWbAreOpened of Sub DieseArbeitsmappe"

End Sub
