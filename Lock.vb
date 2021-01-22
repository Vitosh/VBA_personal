Sub ProtectCellsWithFormulas()
    'lock cells, lock ranges, lock cells with formulas
    Dim wks As Worksheet
    Dim myCell As Range
    
    For Each wks In ThisWorkbook.Worksheets
        If wks.Name = tblForwinCrest.Name Or wks.Name = tblForwinCrestPrefilled.Name Then
            For Each myCell In wks.Range("A1:R102").Cells
                If myCell.MergeArea.Cells.Count = 1 Then
                    If myCell.HasFormula Then
                        myCell.Locked = True
                    Else
                        myCell.Locked = False
                    End If
                End If
            Next myCell
            wks.Protect "vitoshacademy"
        End If
    Next wks
    

End Sub