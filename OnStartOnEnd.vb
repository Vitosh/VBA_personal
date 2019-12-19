Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True

    ActiveWindow.View = xlNormalView
    Application.StatusBar = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    
End Sub

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    
    ActiveWindow.View = xlNormalView
    Application.StatusBar = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False

End Sub
