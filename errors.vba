'Err.Raise 1985, "NAME", "NAME THE CUSTOM ERRROR"
'http://onlinelibrary.wiley.com/doi/10.1002/9781118257616.app3/pdf

Main2_Error:
    
    If Err.Number = [set_standard_error_number] Then
        MsgBox Err.Description & vbCrLf & "Fehler bei Modul " & Err.Source, vbInformation, [set_awaited_error]
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ")", vbInformation, [set_awaited_error_not]
    End If

End Sub
