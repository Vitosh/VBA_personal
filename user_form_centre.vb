Private Sub UserForm_Activate()
    With Me
        .Top = CLng((Application.Height / 2 + Application.Top) - .Height / 2)
        .Left = CLng((Application.Width / 2 + Application.Left) - .Width / 2)
    End With
End sub
