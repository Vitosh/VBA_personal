Public Sub SetWorkedBy()
    
    Set my_cell = tbl_plan.Cells(obj_plan.LastLine, obj_cal.RightColPosition)
    my_cell = "WorkedBy: " & Application.WorksheetFunction.Proper(Environ("username")) & " - " & Format(Now(), "Short Date")
    my_cell.HorizontalAlignment = xlRight
    
End Sub 
