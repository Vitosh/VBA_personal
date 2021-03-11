Public Sub SetPrintArea()

    Dim r_print_range           As Range
    
    Set r_print_range = tbl_plan.Range(Cells(1, 1), Cells(obj_plan.LastLine, obj_cal.RightColPosition))
    
    With tbl_plan.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""Calibri,bold""&25" & "Ankaufsunterlagen"
        
        .PrintArea = r_print_range.Address
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
        
    End With
    
End Sub
