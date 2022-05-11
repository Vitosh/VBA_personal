Sub GetTheMdx()

    Dim pvtTable As PivotTable
    Set pvtTable = tblFoo.PivotTables(1)
    Dim result As String
    result = pvtTable.MDX & "---END"
    Debug.Print result
    
End Sub
