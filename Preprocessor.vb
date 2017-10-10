Option Explicit

#If Win32 Then

    Sub MyTest()
        Debug.Print "32 bits."
    End Sub
    
#ElseIf Win64 Then

    Sub MyTest()
        Debug.Print "64 bits."
        'This should be an error only if it is 64 bits:
        Debug.Print 0 / 0
    End Sub
    
#ElseIf Win16

    Sub MyTest()
        Debug.Print "16 bits."
    End Sub
    
#End If

Sub MyExecutiveMain()
    
    MyTest

End Sub

Sub WhichVersion()

    #If VBA7 Then
        Debug.Print "VBA7"
    #Else
        Debug.Print "NOT VBA7"
    #End If

End Sub
