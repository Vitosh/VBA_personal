Public Sub ColorTheColumn()
    
    Dim l_counter                       As Long
    Dim my_cell                         As Range
    Dim my_cell_find                    As Range
    
    For l_counter = 1 To l_writing_row
        Set my_cell = tbl_output.Cells(l_counter, 1)
        Set my_cell_find = tbl_settings.Range("CN:CN").Find(my_cell, LookIn:=xlValues)
        
        If Not my_cell_find Is Nothing Then
            If my_cell_find.Offset(0, 1) = "bold" Then
                my_cell.Font.Bold = True
            End If
            If my_cell_find.Offset(0, 2) = "red" Then
                my_cell.Font.Color = -16777063
            End If
        End If
        
    Next l_counter
    
End Sub


Public Sub PrintPage()

    Dim Sh                          As Worksheet
    Dim rngPrint                    As Range
    Dim s_reduce_paper_title        As String
    
   On Error GoTo PrintPage_Error
    
    s_reduce_paper_title = "Reduzieren Sie den Papierverbrauch"
    
    Set Sh = ActiveSheet
    Set rngPrint = [input_print_area]
    
    With Sh.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
    Select Case MsgBox("Sind Sie sicher, dass Sie drucken moechten?", vbYesNo Or vbQuestion Or vbDefaultButton1, s_reduce_paper_title)
        Case vbYes
            Select Case MsgBox("Wirklich sicher, dass Sie drucken moechten?", vbYesNo Or vbQuestion Or vbDefaultButton1, s_reduce_paper_title)
                Case vbYes
                rngPrint.PrintOut
        End Select
    End Select

   On Error GoTo 0
   Exit Sub

PrintPage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPage of Modul mod_Drucken"
    
End Sub

Public Sub print_array(my_array As Variant)
    Dim counter As Integer
    
    For counter = LBound(my_array) To UBound(my_array)
        Debug.Print counter & " --> " & my_array(counter)
    Next counter
    
End Sub

Public Sub GenerateSumsOutput(l_lower_row As Long, l_higher_row As Long, l_current_row As Long)

    Dim r_cell              As Range
    Dim l_counter           As Long

    For l_counter = arr_calendar_settings(2) To arr_calendar_settings(3)
        Set r_cell = tbl_output.Cells(l_current_row, l_counter)
        r_cell.FormulaR1C1 = "=SUM(R" & l_higher_row & "C:R" & l_lower_row & "C)"
    Next l_counter

    Set r_cell = Nothing
    
End Sub

Public Sub swap_variables(ByRef value_1, ByRef value_2)
    
    Dim int_tmp                 As Integer
    
    int_tmp = value_1
    value_1 = value_2
    value_2 = int_tmp
    
End Sub

Public Function calculate_years_from_months(total_term) As Long
    
    calculate_years_from_months = total_term \ MONTHS_IN_YEAR
    If total_term Mod MONTHS_IN_YEAR Then calculate_years_from_months = calculate_years_from_months + 1
    
End Function

Public Function letter_col(ByVal col As Long) As String

    letter_col = Split(Cells(1, col).Address, "$")(1)

End Function

Public Function bool_zero_or_empty(ByRef cell As Range, Optional b_is_range = False) As Boolean
    
    If b_is_range Then
        
        For Each current_cell In cell
            If (IsEmpty(current_cell) Or current_cell.Value = 0) Then
                bool_zero_or_empty = True
                Exit Function
            Else
                bool_zero_or_empty = False
            End If
        Next current_cell
        
    Else
        If (IsEmpty(cell) Or cell.Value = 0) Then
            bool_zero_or_empty = True
        Else
            bool_zero_or_empty = False
        End If
    End If

End Function

Public Function change_commas(ByVal myValue As Variant) As String
    
    Dim str_temp As String
    
    str_temp = CStr(myValue)
    change_commas = Replace(str_temp, ",", ".")
    
End Function

Public Sub FormatAsDate(ByRef cell As Range)

    cell.NumberFormat = "[$-407]mmm/ yy;@"
    
End Sub

Public Sub FormatAsPercent(ByRef my_cell As Range)

    my_cell.Style = "Percent"
    my_cell.NumberFormat = "0.00%"

End Sub

Public Sub FormatAsCurrency(ByRef cell As Range, Optional ByVal b_change_0 = False, Optional b_make_gray = True)
    
    Dim b_is_alone          As Boolean
    
    b_is_alone = IIf(cell.Rows.Count + cell.Columns.Count <> 2, False, True)

    If IsNumeric(cell.Value) And Not cell.HasFormula Then
        cell.Value = Round(cell.Value, 2)
    End If

    cell.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

    If b_change_0 Then

        With cell
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .FormatConditions(1).Font.ThemeColor = xlThemeColorDark1
            .FormatConditions(1).Font.TintAndShade = -0.4
        End With
    End If

    If b_is_alone Then
        If b_make_gray And cell.Value = 0 Then
            With cell
                .Cells.Font.Color = RGB(191, 191, 191)
            End With
        End If
    End If

End Sub

Public Function millions_eur(ByVal my_value As Long) As Long
    
    millions_eur = my_value / 1000000

End Function

Public Sub WhiteYourself(ByVal lines As Long, ByRef my_sheet As Worksheet)
    
    Dim str_lines As String
    str_lines = lines & ":" & lines
    
    With my_sheet.Rows(str_lines).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
End Sub

Public Sub WhiteCell(ByRef my_cell As Range)
    
    my_cell.Font.ThemeColor = xlThemeColorDark1
    my_cell.Font.TintAndShade = 0
    
End Sub

Public Sub FormatFontColorToGrey(ByRef cell As Range)

    cell.Font.Color = RGB(128, 128, 128)

End Sub

Public Function sum_range(my_range As Range) As Double

    Dim cell As Range

    sum_range = 0
    For Each cell In my_range
        sum_range = sum_range + cell.Value
    Next

End Function


Public Function make_random(down As Integer, up As Integer)

    make_random = Int((up - down + 1) * Rnd + down)

End Function

Public Function last_row_with_data(ByVal lng_column_number As Long, shCurrent As Variant) As Long
    
    last_row_with_data = shCurrent.Cells(Rows.Count, lng_column_number).End(xlUp).Row
    
End Function

Sub CopyValues(rngSource As Range, rngTarget As Range)
 
    rngTarget.Resize(rngSource.Rows.Count, rngSource.Columns.Count).Value = rngSource.Value
 
End Sub

Public Sub FormatRedAndBold(ByRef my_cell As Range, Optional isBold = True)
    
    my_cell.Font.Color = -16777063
    my_cell.Font.TintAndShade = 0

    If isBold Then my_cell.Font.Bold = True
    
End Sub

Public Function check_if_hidden(r_range As Range) As Boolean

    If r_range.EntireRow.Hidden Or r_range.EntireColumn.Hidden Then
        check_if_hidden = True
    End If

End Function

Function NamedRangeExists(strRangeName As String) As Boolean
    Dim my_range As Range
    
    On Error Resume Next
    
    Set my_range = Range(strRangeName)
    
    If Not my_range Is Nothing Then NamedRangeExists = True
    
    On Error GoTo 0
    
End Function

Public Sub FormatAs_Eur_pro_m2(my_cell As Range)
    
    my_cell.NumberFormat = "#,##0.00 "" € / m²"""

End Sub

Sub change_all_names()
    
    Dim i               As Integer
    Dim s_old           As String
    Dim s_new           As String
    
    For i = 1 To ActiveWorkbook.Names.Count
'        Debug.Print ActiveWorkbook.Names(i).name
'        Debug.Print ActiveWorkbook.Names(i).RefersToR1C1
'        Debug.Print ActiveWorkbook.Names(i)
'
        If InStr(1, ActiveWorkbook.Names(i), "old", vbTextCompare) Then
            s_old = ActiveWorkbook.Names(i).RefersToR1C1
            s_new = Replace(s_old, "old", "")
            Debug.Print s_new
            
            With ActiveWorkbook.Names(ActiveWorkbook.Names(i).name)
                .RefersToR1C1 = s_new

            End With
        End If
    Next i

End Sub
Sub Fixing()
    tbl_Input.img_coat_of_arms.BackColor = RGB(217, 217, 217)
End Sub

Public Sub SetUserNameAndDate()

    [input_calculation_date] = Date
    [input_user_name] = "Erstellt von " & Replace(Application.WorksheetFunction.Proper(Environ("UserName")), ".", ". ")

End Sub


Public Sub SetNamedRanges()

    'start setting named range for ma_purchase_ba
    If NamedRangeExists("ma_purchase_ba") Then ActiveWorkbook.Names("ma_purchase_ba").Delete
    ThisWorkbook.Names.Add name:="ma_purchase_ba", RefersTo:=tbl_output.Cells(8, 3)
    'end   setting named range

End Sub

Public Function locate_bau_beginn(ByVal d_baubeginn As Date) As Long
    
    Dim cell_to_find As Range
    
    Set cell_to_find = Range(tbl_output.Cells(1, 1), tbl_output.Cells(1, arr_calendar_settings(3))).Find(d_baubeginn, LookIn:=xlValues)
    locate_bau_beginn = cell_to_find.Column
    Set cell_to_find = Nothing
    
End Function

Public Function get_last_day_of_month(ByVal my_date As Date) As Date
    get_last_day_of_month = DateSerial(Year(my_date), month(my_date) + 1, 0)
End Function

Public Function get_first_day_of_month(ByVal my_date As Date) As Date
    get_first_day_of_month = DateSerial(Year(my_date), month(my_date), 1)
End Function

Public Function add_months(ByVal my_date As Date, ByVal i_month As Integer) As Date
    add_months = get_last_day_of_month(DateAdd("m", i_month, my_date))
End Function

Public Function add_months_and_get_first_date(ByVal my_date As Date, ByVal i_month As Integer) As Date
    add_months_and_get_first_date = get_first_day_of_month(DateAdd("m", i_month, my_date))
End Function

Public Sub FreezePanesWithoutSelect()

    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    Set ws = Worksheets("master")
    
    Application.Goto ws.Range("E2")
    ActiveWindow.FreezePanes = True
    
    Set ws = Nothing
    
End Sub

Public Function get_column_with_value(ByRef my_cell) As Long

    get_column_with_value = my_cell.End(xlToRight).Column

End Function

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

Public Sub UpdateStatusBar()
    
    Dim i                   As Integer
    Dim s_show              As String

   On Error GoTo UpdateStatusBar_Error

    If int_number_of_subs = 0 Then int_number_of_subs = 1

    int_current_sub = int_current_sub + 1

    s_show = "/\/\>-"
    
    For i = 0 To int_number_of_subs Step 1
        If int_current_sub <> i Then
            s_show = s_show & "~~~"
        Else
            s_show = s_show & "\___/"
        End If
    Next i

    s_show = s_show & "-</\/\"

    Application.StatusBar = s_show
    
   On Error GoTo 0
   Exit Sub

UpdateStatusBar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateStatusBar of Modul mod_StatusBarAndSelection"
    
End Sub

Public Sub SelectMeA1RangeEverywhere()
    
    Dim Sheet As Worksheet

    For Each Sheet In ThisWorkbook.Sheets
        If Sheet.Visible = xlSheetVisible Then
            Sheet.Activate
            Sheet.Cells(1, 1).Select
        End If
    Next Sheet
    
    tbl_paku.Select

   Exit Sub

End Sub

sub WithoutSelectFreezePanes

    Application.Goto tbl_output.Cells(3, 6)
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True

end sub

Function bubble_sort(ByRef TempArray As Variant) As Variant
    Dim Temp            As Variant
    Dim i               As Integer
    Dim NoExchanges     As Integer
    
    ' Loop until no more "exchanges" are made.
    Do
        NoExchanges = True
        
        ' Loop through each element in the array.
        For i = LBound(TempArray) To UBound(TempArray) - 1
        
            ' If the element is greater than the element
            ' following it, exchange the two elements.
            If CLng(TempArray(i)) > CLng(TempArray(i + 1)) Then
                NoExchanges = False
                Temp = TempArray(i)
                TempArray(i) = TempArray(i + 1)
                TempArray(i + 1) = Temp
            End If
        Next i
    
    Loop While Not (NoExchanges)
    bubble_sort = TempArray
End Function

Public Function sum_array(my_array As Variant) As Double
    'For unknown reasons, WorksheetFunction.sum(my_array) does not work always,
    'when we sum currency, integer and double...
    
    Dim l_counter           As Long
    
    For l_counter = LBound(my_array) To UBound(my_array)
        sum_array = sum_array + my_array(l_counter)
    Next l_counter
    
End Function

Public Function b_value_in_array(my_value As Variant, my_array As Variant) As Boolean

    Dim l_counter
    
    For l_counter = LBound(my_array) To UBound(my_array)
        my_array(l_counter) = CStr(my_array(l_counter))
    Next l_counter

    b_value_in_array = Not IsError(Application.Match(CStr(my_value), my_array, 0))
    
End Function
