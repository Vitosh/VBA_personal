used this way:
If IsUserFormLoaded("frmPlanerkostenberechnung") Then Unload frmPlanerkostenberechnung

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit Function
        End If
    Next
End Function
