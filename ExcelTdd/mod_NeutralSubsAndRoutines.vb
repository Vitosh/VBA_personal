Option Explicit

Public Sub Increment(ByRef value_to_increment, Optional l_plus As Double = 1) 'optional value type changed to_double
    
    value_to_increment = value_to_increment + l_plus
    
End Sub


Public Function GetDateAndTime() As String

    GetDateAndTime = Format(DateValue(Date), "dd-mm-yyyy") & " " & Time

End Function

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    ActiveWindow.View = xlNormalView

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

Public Function codify_time(Optional b_make_str As Boolean = False) As String

    If SET_IN_PRODUCTION Then On Error GoTo codify_Error
    
    Dim dbl_01                  As Variant
    Dim dbl_02                  As Variant
    Dim dbl_now                 As Double
    
    dbl_now = Round(Now(), 8)
    
    dbl_01 = Split(CStr(dbl_now), ",")(0)
    dbl_02 = Split(CStr(dbl_now), ",")(1)
    
    codify_time = Hex(dbl_01) & "_" & Hex(dbl_02)
    
    If b_make_str Then codify_time = "\" & codify_time & ".txt"
    
    On Error GoTo 0
    Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function
