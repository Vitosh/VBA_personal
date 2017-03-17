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

Public Sub HideNeeded()
    
    Dim var_Sheet                   As Variant
    
    Dim arr_visible_sheets          As Variant
    Dim arr_hidden_sheets           As Variant
    
    Call OnStart
     
    arr_visible_sheets = Array(tbl_Input)
    arr_hidden_sheets = Array(tbl_1, tbl_2, tbl_3)
    
    For Each var_Sheet In arr_visible_sheets
        var_Sheet.Visible = xlSheetVisible
    Next var_Sheet
    
    For Each var_Sheet In arr_hidden_sheets
        var_Sheet.Visible = xlSheetVeryHidden
    Next var_Sheet
   
    Call OnEnd
    
End Sub


Public Sub UnhideAll()
        
    Dim Sheet As Worksheet
    
    For Each Sheet In ThisWorkbook.Worksheets
       ' If Sheet.Visible = Not xlSheetVisible Then Sheet.Visible = xlSheetVisible
       Sheet.Visible = xlSheetVisible
    Next Sheet
    
    Call UnprotectAll
    
End Sub
