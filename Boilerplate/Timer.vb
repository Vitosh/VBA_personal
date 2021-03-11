Public myTime As Double

Sub SetPublicConstants()
    
    myTime = Timer
    
End Sub

'Option 1
Public Sub ShowMsgBoxReady1()

    Debug.Print "Ready" & vbCrLf & Round(Timer - myTime) & " Seconds!", vbInformation, "Information"

End Sub

'Option 2
Public Sub ShowMsgBoxReady2()
    
    Dim timeAsText As String
    
    timeAsText = (Timer - myTime) \ 60 & ":" & (Timer - myTime) Mod 60
    Debug.Print "Ready!" & vbCrLf & timeAsText, vbInformation, "Information"

End Sub

'Option 3
Public Sub ShowMsgBoxReady3()
    
    Dim timeAsText As String
    Dim result As Long
    
    result = Timer - myTime

    timeAsText = result \ 60 & ":" & IIf(Len(CStr(result Mod 60)) < 2, "0", "") & (result) Mod 60
    Debug.Print "Ready!" & vbCrLf & timeAsText, vbInformation, "Information"

End Sub



