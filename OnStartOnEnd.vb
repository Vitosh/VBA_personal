Public Sub OnStart()
    
    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Application.EnableEvents = False
    Application.DisplayAlerts = False

End Sub

Public Sub OnEnd()

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.AskToUpdateLinks = True
    
End Sub
