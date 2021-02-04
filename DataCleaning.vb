Sub AddValueIfEmptyInCellBelow()

    Dim myCell As Range
    Dim str As String
    
    For Each myCell In Selection
        If Len(Trim(myCell)) = 0 Then
            myCell = str
        Else
            str = myCell
        End If
    Next myCell

End Sub
