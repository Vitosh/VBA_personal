Public Sub IgnoreCellErrors()
    
    Dim rngCell     As Range
    Dim cnt         As Long
    
    For Each rngCell In ActiveSheet.UsedRange
        For cnt = 1 To 8
            rngCell.Errors(cnt).Ignore = True
        Next cnt
    Next rngCell

End Sub
