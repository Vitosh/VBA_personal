Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    On Error GoTo Workbook_BeforeClose_Error
    
    If Not SET_IN_PRODUCTION Then
        MsgBox "SET_IN_PRODUCTION"
        On Error GoTo 0
        Cancel = True
    End If
    
    Cancel = False
    
    ThisWorkbook.Save

    Application.DisplayAlerts = False
    HideNeededWorksheets
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
    Application.DisplayAlerts = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    'ActiveSheet.PageSetup.BlackAndWhite = True
    Me.Save

    EnableMySaves

    On Error GoTo 0
    Exit Sub

Workbook_BeforeClose_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_BeforeClose"

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
        
    If Not SET_IN_PRODUCTION Then
        MsgBox "SET_IN_PRODUCTION", vbInformation, CON_STR_APP_NAME
        Cancel = True
    End If
    
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)

    If Not tblSettings.Visible Then
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

    HideNeededWorksheets
    'Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", false)"
    'Application.DisplayFormulaBar = False

    If Not IsValueInArray(Environ("username"), ADMINS, True) Then
        Application.OnKey "%{F11}", "DisabledCombination"
    End If

    DisableShortcutsAndSaves

    If ThisWorkbook.Date1904 Then
        MsgBox CON_STR_1904, vbInformation, CON_STR_APP_NAME
    End If

    Application.WindowState = xlMaximized

    CheckHowManyWbAreOpened

    On Error GoTo 0
    Exit Sub

Workbook_Open_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Workbook_Open"
    Me.Save
    ThisWorkbook.Close

End Sub
