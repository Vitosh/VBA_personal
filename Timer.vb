
Public dbl_timer                        As Double

'On Start:
Sub SetPublicConstants()
    
    dbl_timer = timer
    
End Sub

'At the end
Public Sub ShowMsgBoxReady()

    MsgBox "Fertig" & vbCrLf & Round(timer - dbl_timer) & "Sekunden!", vbInformation, "Nachricht"

End Sub

