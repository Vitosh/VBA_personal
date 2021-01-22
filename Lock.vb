'lock cells, lock ranges, lock cells with formulas
Sub ProtectCellsWithFormulas()
   
    Dim wks As Worksheet
    Dim myCell As Range
    
    For Each wks In ThisWorkbook.Worksheets
        With wks
            If .Name = tblForwinCrest.Name Or .Name = tblForwinCrestPrefilled.Name Then
                .Unprotect "v"
                For Each myCell In wks.Range("A1:R102").Cells
                    If myCell.MergeArea.Cells.Count = 1 Then
                        If myCell.HasFormula Then
                            myCell.Locked = True
                        Else
                            myCell.Locked = False
                        End If
                    End If
                Next myCell
                .EnableOutlining = True
                .Protect "v", contents:=True, userinterfaceonly:=True
            End If
        End With
    Next wks
    

End Sub
