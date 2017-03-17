
Sub ctrl_plus_u()

' Tastenkombination: Strg + U

    If Selection.Font.Underline = xlUnderlineStyleSingle Then
        Selection.Font.Underline = xlUnderlineStyleNone
    Else
        Selection.Font.Underline = xlUnderlineStyleSingle
    End If

End Sub

Sub ctrl_plus_b()

' Tastenkombination: Strg + B
    
    Selection.Font.Bold = Not Selection.Font.Bold

End Sub

Public Sub ctrl_plus_i()

' Tastenkombination: Strg + I
    
    Selection.Font.Italic = Not Selection.Font.Italic

End Sub

Sub ctrl_plus_d()

' Tastenkombination: Strg + D
    
    Selection.FillDown

End Sub
