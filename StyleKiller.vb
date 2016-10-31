Option Explicit

Sub StyleKiller()

    Dim my_style                As Style

    For Each my_style In ThisWorkbook.Styles
        If Not my_style.BuiltIn Then
            Debug.Print my_style.Name
            my_style.Delete
        End If
    Next my_style

End Sub

'FANCY ONE:
'**************************************************************************************
Sub RemoveTheStyles()

    Dim style               As style
    Dim l_counter           As Long
    Dim l_total_number      As Long

    On Error Resume Next

    l_total_number = ActiveWorkbook.Styles.Count
    Application.ScreenUpdating = False

    For l_counter = l_total_number To 1 Step -1
    
        Set style = ActiveWorkbook.Styles(l_counter)
        
        If (l_counter Mod 500 = 0) Then
            DoEvents
            Application.StatusBar = "Deleting " & l_total_number - l_counter + 1 & " of " & l_total_number & " " & style.Name
        End If
        
        If Not style.BuiltIn Then style.Delete

    Next l_counter

    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print "READY!"
    
    On Error GoTo 0
End Sub
