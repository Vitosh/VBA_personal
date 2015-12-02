
Public dbl_timer                        As Double

'On Start:
'[set_execution_time] = "Durchführungszeit: "

Sub SetPublicConstants()
    
    dbl_timer = timer
    
End Sub

'At the end
Public Sub ShowMsgBoxReady()

    MsgBox "Fertig" & vbCrLf & Round(timer - dbl_timer) & "Sekunden!", vbInformation, "Nachricht"

End Sub

Public Sub ShowMsgBoxReady()
    
    Dim str_time As String
    
    str_time = [set_execution_time] & (timer - dbl_timer) \ 60 & ":" & (timer - dbl_timer) Mod 60
    MsgBox "Fertig!" & vbCrLf & str_time, vbInformation, "Nachricht"

End Sub


