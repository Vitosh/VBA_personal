Option Explicit
Option Private Module

Public Sub Tdd()

    Dim lngTestsTotalExpected               As Long

'    Select Case MsgBox("The TDD is probably long.", vbYesNo, "Sure?")
'        Case vbNo
'            Exit Sub
'    End Select
    
    SET_IN_PRODUCTION = False
    
    lngTestsTotalExpected = 999 'PLACEHOLDER_VALUE
    Debug.Print "Test report from " & Environ("Username") & vbCrLf & "START: " & GetDateAndTime & vbCrLf & _
                    lngTestsTotalExpected & " expected." & vbCrLf
    Call OnStart
    Worksheets(1).Select
    
    STR_ERROR_REPORT = "Test report from " & Environ("Username") & vbCrLf & "START: " & GetDateAndTime & vbCrLf & _
                    lngTestsTotalExpected & " expected." & vbCrLf & vbCrLf

    LNG_TOTAL_TESTS = 0

    Call Tdd_01


    STR_ERROR_REPORT = STR_ERROR_REPORT & vbCrLf & "Tests expected: " & lngTestsTotalExpected & vbCrLf & _
                        "Total Tests:" & LNG_TOTAL_TESTS & vbCrLf & "END: " & GetDateAndTime

    [SET_IN_PRODUCTION] = True
    Debug.Print "Tests expected: " & lngTestsTotalExpected
    Debug.Print "Total Tests:" & vbCrLf & LNG_TOTAL_TESTS & vbCrLf & "END: " & GetDateAndTime

    Call CreateLogFile
    Call OnEnd

    STR_ERROR_REPORT = ""

End Sub


