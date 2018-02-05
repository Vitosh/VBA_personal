Public Sub removeSpaceInString()

    Dim myCell As Range
        
    For Each myCell In Selection
        myCell = Trim(myCell)
        myCell = Replace(myCell, vbTab, "")
        myCell = Replace(myCell, " ", "")
        myCell = Replace(myCell, Chr(160), "")
    Next myCell

End Sub
