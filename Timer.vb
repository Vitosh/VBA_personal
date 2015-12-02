
Public dbl_timer                        As Double

'On Start:
'[set_execution_time] = "Durchführungszeit: "

Sub SetPublicConstants()
    
    dbl_timer = timer
    
End Sub

'At the end
'Option 1
Public Sub ShowMsgBoxReady()

    MsgBox "Fertig" & vbCrLf & Round(timer - dbl_timer) & "Sekunden!", vbInformation, "Nachricht"

End Sub

'Option 2
Public Sub ShowMsgBoxReady()
    
    Dim str_time As String
    
    str_time = [set_execution_time] & (timer - dbl_timer) \ 60 & ":" & (timer - dbl_timer) Mod 60
    MsgBox "Fertig!" & vbCrLf & str_time, vbInformation, "Nachricht"

End Sub

'Option 3
Public Sub ShowMsgBoxReady()
    
    Dim str_time                As String
    Dim lng_result              As Long
    
    lng_result = Timer - dbl_timer
      
    
    str_time = [set_execution_time] & lng_result \ 60 & ":" & IIf(Len(CStr(lng_result Mod 60)) < 2, "0", "") & (lng_result) Mod 60
    MsgBox "Fertig!" & vbCrLf & str_time, vbInformation, "Nachricht"

End Sub


