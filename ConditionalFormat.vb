Sub ListAllConditionalFormat()
    
    Dim cf      As FormatCondition
    Dim ws      As Worksheet
    Dim l       As Long
    Dim rngCell As Range
    
    On Error Resume Next
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    tblPrint.Cells.Clear
    
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
        
        For Each cf In ws.Cells.FormatConditions
            l = 1 + l
            
            With tblPrint
                
                Set rngCell = Cells(l, 1)
                rngCell = cf.AppliesTo.Address
                rngCell.Offset(0, 1) = cf.Type
                rngCell.Offset(0, 2) = "'" & cf.Formula1
                rngCell.Offset(0, 3) = cf.Interior.Color
                rngCell.Offset(0, 4) = cf.Font.Name
                rngCell.Offset(0, 5) = ws.Name
                rngCell.Offset(0, 6) = "'" & cf.AppliesTo.AddressLocal
                rngCell.Offset(0, 7) = "'" & cf.Formula2
                
                
            End With
        Next cf
    Next ws
    Debug.Print "END!"
    
End Sub
