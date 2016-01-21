Option Explicit

Public Sub main()
    
    On Error GoTo main_Error
    
    Call OnStart
    
    Call ClearWritingPlace
    Call WriteMonthsAbove
    Call GenerateValuesInside
    Call GenerateSumsAtTheEnd
    Call BorderMe(ThisWorkbook.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(L_ROW_WITH_DATES, L_FIRST_COLUMN_TO_WRITE), ThisWorkbook.Sheets(1).Cells(obj_cal.LastRow, obj_cal.RightestColumn)))
    Call AutoFitAndMessageBox
    Call SetObjectsToNothing
    
    Call OnEnd
    
    On Error GoTo 0
    Exit Sub
    
main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main of Module mod_main"
    Call SetObjectsToNothing
    Call OnEnd
    
End Sub

Public Sub SetObjectsToNothing()

    Set r_range_4_dates = Nothing
    Set obj_cal = Nothing

End Sub

Public Sub AutoFitAndMessageBox()
    
    Range(ThisWorkbook.Sheets(1).Cells(L_ROW_WITH_DATES, L_FIRST_COLUMN_TO_WRITE), ThisWorkbook.Sheets(1).Cells(obj_cal.LastRow + 2, obj_cal.RightestColumn)).Columns.AutoFit
    MsgBox STR_FERTIG, vbInformation, STR_SCHADENSERSATZ

End Sub

Public Sub GenerateValuesInside()
    
    Dim l_counter_row               As Long
    Dim l_counter_col               As Long
    
    For l_counter_row = L_STARTING_ROW To obj_cal.LastRow
        For l_counter_col = L_FIRST_COLUMN_TO_WRITE To obj_cal.RightestColumn
            Call GenerateFormula(l_counter_row, l_counter_col)
        Next l_counter_col
    Next l_counter_row
    
End Sub

Public Sub WriteMonthsAbove()

    For l_counter = 0 To obj_cal.CalendarLength - 1
    
        Set my_cell = ThisWorkbook.Sheets(1).Cells(L_STARTING_ROW - 1, L_FIRST_COLUMN_TO_WRITE + l_counter)
        my_cell = add_months(obj_cal.FirstMonth, l_counter)
        Call FormatMyCell(my_cell, False, True, True, True)
        
    Next l_counter
    
End Sub

Public Function last_row_with_data(ByVal lng_column_number As Long, shCurrent As Variant) As Long
    
    last_row_with_data = shCurrent.Cells(Rows.Count, lng_column_number).End(xlUp).Row

End Function

Public Function add_months(my_date As Date, l_month As Long) As Date
    
    add_months = get_last_day_of_month(DateAdd("m", l_month, my_date))

End Function

Public Function get_last_day_of_month(ByVal my_date As Date) As Date
    
    get_last_day_of_month = DateSerial(Year(my_date), Month(my_date) + 1, 0)

End Function

Public Sub ClearWritingPlace()

    Set obj_cal = New cls_calendar
    
    obj_cal.LastRow = last_row_with_data(1, ThisWorkbook.Sheets(1))
    Set r_range_4_dates = ThisWorkbook.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(L_RATE6_VERTRAG_COL, L_STARTING_ROW), ThisWorkbook.Sheets(1).Cells(obj_cal.LastRow, L_RATE5PR_TERMIN_COL))
    
    ThisWorkbook.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(1, L_FIRST_COLUMN_TO_WRITE - 1), ThisWorkbook.Sheets(1).Cells(Rows.Count, Columns.Count)).Clear
    obj_cal.FirstMonth = Application.WorksheetFunction.Min(r_range_4_dates)
    obj_cal.LastMonth = Application.WorksheetFunction.Max(r_range_4_dates)
    obj_cal.CalendarLength = DateDiff("m", obj_cal.FirstMonth, obj_cal.LastMonth)
    obj_cal.RightestColumn = L_FIRST_COLUMN_TO_WRITE + obj_cal.CalendarLength - 1

End Sub


Public Sub GenerateFormula(l_row, l_col)

    Dim date_date_above     As Date
    Dim my_cell             As Range
    Dim l_count_garages     As Long
    Dim b_has_garage        As Boolean: b_has_garage = False
    
    If WorksheetFunction.CountA(Cells(l_row, L_RATE6_VERTRAG_COL)) = 0 Then Exit Sub
    
    dbl_eur_m2 = ThisWorkbook.Sheets(1).Cells(2, 18)
    dbl_eur_garage = ThisWorkbook.Sheets(1).Cells(2, 19)
    
    date_date_above = ThisWorkbook.Sheets(1).Cells(L_ROW_WITH_DATES, l_col)
    Set my_cell = Cells(l_row, l_col)
    
    If Cells(l_row, L_RATE6_VERTRAG_COL) < get_last_day_of_month(Cells(l_row, L_RATE6_TERMIN_COL)) Then
        If date_date_above > Cells(l_row, L_RATE6_VERTRAG_COL) And date_date_above <= get_last_day_of_month(Cells(l_row, L_RATE6_TERMIN_COL)) Then
            my_cell = dbl_eur_m2 * Cells(l_row, 15)
        End If
    End If
    
    On Error Resume Next 'do not do this at home...
    If CLng(Cells(l_row, 3)) > 0 Then b_has_garage = True
    On Error GoTo 0
    
    If Cells(l_row, L_RATE5PR_VERTRAG_COL) < Cells(l_row, L_RATE5PR_TERMIN_COL) And _
        b_has_garage Then
        
        If date_date_above > get_last_day_of_month(Cells(l_row, L_RATE5PR_VERTRAG_COL)) And _
        date_date_above <= get_last_day_of_month(Cells(l_row, L_RATE5PR_TERMIN_COL)) Then
            
            l_count_garages = find_in_string_times(Cells(my_cell.Row, 3)) + 1
        
            my_cell = my_cell + l_count_garages * dbl_eur_garage
        End If
    End If
    
    If my_cell > 0 Then Call FormatMyCell(my_cell, True, False, False, True)

End Sub

Public Function find_in_string_times(my_cell As Range, Optional ch_char As String = "+") As Long

    find_in_string_times = UBound(Split(my_cell, ch_char))

End Function

Public Sub FormatMyCell(ByRef my_cell As Range, Optional b_as_currency As Boolean = False, _
                                                Optional b_as_date As Boolean = False, _
                                                Optional b_as_dark As Boolean = False, _
                                                Optional b_as_din As Boolean = False)
                                                
    If b_as_currency Then
        my_cell.NumberFormat = "#,##0.00 $"
    End If
    
    If b_as_date Then
        my_cell.NumberFormat = "[$-407]mmm/ yy;@"
    End If
    
    If b_as_dark Then
        my_cell.Interior.ThemeColor = xlThemeColorDark1
        my_cell.Interior.TintAndShade = -0.249946592608417
    End If
    
    If b_as_din Then
        my_cell.Font.Name = "DIN-Light"
    End If

End Sub

Public Sub OnStart()

    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Application.EnableEvents = False

End Sub

Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
End Sub

Public Sub GenerateSumsAtTheEnd()

    Dim l_counter           As Long
    Dim my_cell             As Range
    
    obj_cal.LastRow = obj_cal.LastRow + 1
    
    For l_counter = 0 To obj_cal.CalendarLength - 1
        Set my_cell = ThisWorkbook.Sheets(1).Cells(obj_cal.LastRow, L_FIRST_COLUMN_TO_WRITE + l_counter)
        my_cell.FormulaR1C1 = "=SUM(R6C:R" & obj_cal.LastRow - 1 & "C)"
        
        Call FormatMyCell(my_cell, True, False, True, True)

    Next l_counter
    
    Set my_cell = Cells(obj_cal.LastRow + 1, L_FIRST_COLUMN_TO_WRITE)
    my_cell.FormulaR1C1 = "=SUM(R[-1]C:R[-1]C" & L_FIRST_COLUMN_TO_WRITE + obj_cal.CalendarLength - 1 & ")"
    
    Call FormatMyCell(my_cell, True, False, True, True)
    
End Sub

Public Sub BorderMe(my_range)

    Dim l_counter   As Long
    For l_counter = 7 To 10 '7 to 10 are the magic numbers for xlEdgeLeft etc
        With my_range.Borders(l_counter)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Next l_counter
End Sub


Public Sub AddAButton()
    Dim my_btn          As Button
    Dim my_range        As Range
    
    Set my_range = Sheets(1).Cells(1, 19)
    Set my_btn = Sheets(1).Buttons.Add(my_range.Left, my_range.Top, my_range.Width, my_range.Height)
      
    my_btn.OnAction = "main"
    my_btn.Caption = "Laufen"
    my_btn.Name = "created_by_macro"
 
End Sub
