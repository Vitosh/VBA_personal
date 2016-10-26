Sub StyleKiller()

    Dim my_style                As Style

    For Each my_style In ThisWorkbook.Styles
        If Not my_style.BuiltIn Then
            Debug.Print my_style.Name
            my_style.Delete
        End If
    Next my_style

End Sub
