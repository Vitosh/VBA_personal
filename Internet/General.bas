Attribute VB_Name = "General"
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    
    Application.StatusBar = False
    
End Sub

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    
    ActiveWindow.View = xlNormalView

End Sub

Public Sub LogMe(ParamArray arg() As Variant)

    Debug.Print Join(arg, "--")
    
End Sub

Public Sub PrintMeUsefulFormula()

    Dim strFormula  As String
    Dim strParenth  As String

    strParenth = """"

    strFormula = Selection.FormulaR1C1
    
    strFormula = Replace(strFormula, """", """""")

    strFormula = strParenth & strFormula & strParenth
    Debug.Print strFormula
    
End Sub

Public Sub WaitSomeMilliseconds(Optional Milliseconds As Long = 1000)
    Sleep Milliseconds
End Sub
