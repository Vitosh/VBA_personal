'https://msdn.microsoft.com/en-us/library/office/ff197461.aspx
    
Public Sub EnableControls()

    Application.OnKey "^{F8}", "F8_CtrlMacro"
    Application.OnKey "%{F8}", "F8_AltMacro"
    Application.OnKey "+{F8}", "F8_ShiftMacro"
    Application.OnKey "{F8}", "F8_OnlyMacro"
    
End Sub

Public Sub DisableControls()

    Application.OnKey "^{F8}", ""
    Application.OnKey "%{F8}", ""
    Application.OnKey "+{F8}", ""
    Application.OnKey "{F8}", ""
    
End Sub

Public Sub F8_CtrlMacro()
    Debug.Print "F8 with Ctrl"
End Sub

Public Sub F8_AltMacro()
    Debug.Print "F8 with Alt"
End Sub

Public Sub F8_ShiftMacro()
    Debug.Print "F8 with Shift"
End Sub

Public Sub F8_OnlyMacro()
    Debug.Print "F8 Only"
End Sub
