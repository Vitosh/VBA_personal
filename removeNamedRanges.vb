Public Sub RemoveNamedRanges()
    
    Dim nName                   As Name
    Dim strNameReserved         As String
    
    On Error Resume Next
    
    strNameReserved = "set_in_production"
    
    For Each nName In Names
        If nName.Name <> strNameReserved And Left(nName.Name, 1) <> "_" Then
            Debug.Print nName.Name
            nName.Delete
        End If
    Next nName
    
    On Error GoTo 0
    
End Sub
