'Info for form VBEditor, VBE Form

Public Sub InfoForForm()

    Dim cCont As Control

    For Each cCont In Controls
        Debug.Print TypeName(cCont)
        Debug.Print cCont.name
    Next cCont
    
End Sub
