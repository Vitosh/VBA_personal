Option Explicit

'---------------------------------------------------------------------------------------
' Method : ErrorsNumber
' Author : v.doynov
' Date   : 06.04.2017
' Purpose: Model to see how excel calculates floating point numbers.
'---------------------------------------------------------------------------------------
' 0/2 + 0/4 + 0/8 + 1/16 + 1/32 +0/64 + 0/128 + 1/256 + 0/256 +1/512 +0/1024 + 0/2048
' 0,099609375 
'---------------------------------------------------------------------------------------

Public Sub ErrorsNumber()

    Const DIFF_DEFAULT = 0.1
    ThisWorkbook.PrecisionAsDisplayed = False
    Dim lngEndNumber        As Long: lngEndNumber = 30

    Dim dblStarter          As Double
    Dim dblEnder            As Double
    Dim dblDiff             As Double

    Dim lngCounter          As Long
    Dim lngCounter2         As Long
    Dim lngRow              As Long

    Dim dblResult           As Double
    Dim lngCountErrors      As Long
    Dim myCell              As Range

    If lngEndNumber > 10000 Then Debug.Print lngEndNumber & "is too big, it takes too much time!": Exit Sub

    Call OnStart
    Cells.Clear

    For lngCounter = 0 To lngEndNumber
        dblDiff = DIFF_DEFAULT

        For lngCounter2 = 0 To 9
            dblDiff = DIFF_DEFAULT * lngCounter2

            lngRow = lngRow + 1
            Set myCell = Cells(lngRow, 1)

            dblStarter = lngCounter + dblDiff
            dblEnder = lngCounter + dblDiff + DIFF_DEFAULT
            dblResult = dblStarter - dblEnder

            myCell = dblStarter
            myCell.Offset(0, 1) = dblEnder
            myCell.Offset(0, 2).FormulaR1C1 = "=RC[-1]-RC[-2]"
            myCell.Offset(0, 2).NumberFormat = "0.00000000000000000"
            myCell.Offset(0, 3).FormulaR1C1 = "=IF(RC[-1]=0.1,"""",""X"")"
            
        Next lngCounter2
        
        If lngCounter Mod 100 = 0 Then Debug.Print lngCounter
        
    Next lngCounter

    With Range("E1")
        .FormulaR1C1 = "=COUNTIF(C[-1],""X"")/" & lngEndNumber * 10
        .NumberFormat = "0.0000%"
    End With

    Columns.AutoFit
    Debug.Print "READY!"

    Call OnEnd

End Sub

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

