Public Function change_commas(ByVal myValue As Variant) As String
    
    Dim str_temp As String
    
    str_temp = CStr(myValue)
    change_commas = Replace(str_temp, ",", ".")
    
End Function

Public Function bubble_sort(ByRef TempArray As Variant) As Variant
    Dim Temp            As Variant
    Dim i               As Long
    Dim NoExchanges     As Long
    
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

   On Error GoTo 0
   Exit Function
   
End Function

Public Function get_last_day_of_month(ByVal my_date As Date) As Date

    get_last_day_of_month = DateSerial(Year(my_date), Month(my_date) + 1, 0)
    
End Function

Public Function get_first_day_of_month(ByVal my_date As Date) As Date
    
    get_first_day_of_month = DateSerial(Year(my_date), Month(my_date), 1)

End Function

Public Function add_months(ByVal my_date As Date, ByVal i_month As Long) As Date
    
    add_months = get_last_day_of_month(DateAdd("m", i_month, my_date))

End Function

Public Function add_months_and_get_first_date(ByVal my_date As Date, ByVal i_month As Long) As Date

    add_months_and_get_first_date = get_first_day_of_month(DateAdd("m", i_month, my_date))

End Function

Public Function calculate_years_from_months(total_term) As Long
    
    calculate_years_from_months = total_term \ MONTHS_IN_YEAR
    If total_term Mod MONTHS_IN_YEAR Then calculate_years_from_months = calculate_years_from_months + 1
    
End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
    
    On Error Resume Next
    
        IsArrayAllocated = IsArray(Arr) And Not IsError(LBound(Arr, 1)) And LBound(Arr, 1) <= UBound(Arr, 1)
    
    On Error GoTo 0

End Function

Public Sub print_array(ByRef my_array As Variant)
    Dim counter As Long
    
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

Public Sub FormatAsDate(ByRef cell As Range)

    cell.NumberFormat = "[$-407]mmm/ yy;@"
    
End Sub

Public Sub FormatAsPercent(ByRef my_cell As Range, Optional l_numbers = 2)

    If l_numbers = 3 Then
        my_cell.NumberFormat = "0.000%"
    Else
        my_cell.NumberFormat = "0.00%"
    End If

End Sub

Public Sub FormatAsCurrency(ByRef cell As Range, Optional ByVal b_change_0 = False, Optional b_make_gray = True, Optional b_make_round = True)
    
    Dim b_is_alone          As Boolean
    
    b_is_alone = IIf(cell.Rows.Count + cell.Columns.Count <> 2, False, True)

    If IsNumeric(cell.Value) And (Not cell.HasFormula) Then
        cell.Value = Round(cell.Value, 2)
    End If
    
    If b_make_round Then
        cell.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Else
        cell.NumberFormat = "$#,##0.00_);($#,##0.00)"
    End If
    
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

Public Sub FormatAs_Eur_pro_m2(my_cell As Range)
    
    my_cell.NumberFormat = "#,##0.00 "" € / m²"""

End Sub

Public Sub FormatRedAndBold(ByRef my_cell As Range, Optional isBold = True)
    
    my_cell.Font.Color = -16777063
    my_cell.Font.TintAndShade = 0

    If isBold Then my_cell.Font.Bold = True
    
End Sub

Public Function millions_eur(ByVal my_value As Long) As Long
    
    millions_eur = my_value / 1000000

End Function

Public Sub WhiteYourself(ByVal lines As Long, ByRef my_sheet As Worksheet)
    
    Dim str_lines                       As String
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

Public Function make_random(down As Long, up As Long)

    make_random = Int((up - down + 1) * Rnd + down)

End Function

Public Function last_row_with_data(ByVal lng_column_number As Long, shCurrent As Variant) As Long
    
    last_row_with_data = shCurrent.Cells(Rows.Count, lng_column_number).End(xlUp).row
    
End Function

Sub CopyValues(rngSource As Range, rngTarget As Range)
 
    rngTarget.Resize(rngSource.Rows.Count, rngSource.Columns.Count).Value = rngSource.Value
 
End Sub

Public Function check_if_hidden(r_range As Range) As Boolean

    If r_range.EntireRow.Hidden Or r_range.EntireColumn.Hidden Then
        check_if_hidden = True
    End If

End Function

Function last_row(Optional str_sheet As String, Optional column_to_check As Long = 1) As Long
    
    Dim shSheet             As Worksheet
    
    If str_sheet = vbNullString Then
        Set shSheet = ActiveSheet
    Else
        Set shSheet = Worksheets(str_sheet)
    End If
    
    last_row = shSheet.Cells(shSheet.Rows.Count, column_to_check).End(xlUp).row

End Function

Function last_column(Optional str_sheet As String, Optional row_to_check As Long = 1) As Long

    Dim shSheet  As Worksheet
    
    If str_sheet = vbNullString Then
        Set shSheet = ActiveSheet
    Else
        Set shSheet = Worksheets(str_sheet)
    End If
    
    last_column = shSheet.Cells(row_to_check, shSheet.Columns.Count).End(xlToLeft).Column
    
End Function

Public Function letter_col(ByVal col As Long) As String

    letter_col = Split(Cells(1, col).Address, "$")(1)

End Function

Public Function b_value_in_array(my_value As Variant, my_array As Variant, Optional b_is_string As Boolean = False) As Boolean

    Dim l_counter

    If b_is_string Then
        my_array = Split(my_array, ":")
    End If

    For l_counter = LBound(my_array) To UBound(my_array)
        my_array(l_counter) = CStr(my_array(l_counter))
    Next l_counter

    b_value_in_array = Not IsError(Application.Match(CStr(my_value), my_array, 0))
    
End Function

Public Sub DrawBordersAroundRange(b_remove As Boolean)

    If b_remove Then

        [set_format].Copy
        [input_all_ba].PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        
        'make the last month white for austria
        If tbl_Input.opt_os Then
            For Each current_cell In [input_construction_time]
                tbl_Input.Cells(current_cell.row + 8, 12).Font.Color = vbWhite
            Next current_cell
        End If
        
    Else
        [set_format_without_borders].Copy
        [input_all_ba].PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = xlNone
    End If

End Sub

Public Sub UnhideAll()
        
    Dim Sheet As Worksheet
    
    For Each Sheet In ThisWorkbook.Worksheets
       ' If Sheet.Visible = Not xlSheetVisible Then Sheet.Visible = xlSheetVisible
       Sheet.Visible = xlSheetVisible
    Next Sheet
    
    Call UnprotectAll
    
End Sub

Public Sub UnprotectAll()

    Dim i As Long
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
        ActiveWorkbook.Worksheets(i).Unprotect Password:=s_CONST
    Next
    
End Sub

Public Sub HideNeeded()
    
    Dim var_Sheet                   As Variant
    
    Dim arr_visible_sheets          As Variant
    Dim arr_hidden_sheets           As Variant
    
    Call OnStart
     
    arr_visible_sheets = Array(tbl_Input)
    arr_hidden_sheets = Array(tbl_output, tbl_calendar, tbl_log, tbl_settings, tbl_results, tbl_settings_bau)
    
    For Each var_Sheet In arr_visible_sheets
        var_Sheet.Visible = xlSheetVisible
    Next var_Sheet
    
    For Each var_Sheet In arr_hidden_sheets
        var_Sheet.Visible = xlSheetVeryHidden
    Next var_Sheet
   
    Call OnEnd
    
End Sub

Public Sub add_comment_to_selection(my_comment As Range)
    Dim b As Boolean
    b = True
    For Each current_cell In Selection
        If b Then
            current_cell.ClearComments
            current_cell.AddComment my_comment.Text
            current_cell.Comment.Visible = False
            current_cell.Comment.Shape.ScaleWidth 4, msoFalse, msoScaleFromTopLeft
            current_cell.Comment.Shape.ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft
        End If
        b = Not b
    Next current_cell
End Sub

Public Sub delete_comment_in_selection()
    For Each current_cell In Selection
        current_cell.ClearComments
    Next current_cell
End Sub

Sub DeleteDrawingObjects()

    Dim l_counter           As Long
    
    For l_counter = tbl_Input.DrawingObjects().Count To 1 Step -1
        'Debug.Print tbl_Input.DrawingObjects(l_counter).name
        If Left(tbl_Input.DrawingObjects(l_counter).Name, 7) = "TextBox" Then
            tbl_Input.DrawingObjects(l_counter).Delete
        End If
    Next l_counter

End Sub

Sub CoverRange(ByRef R As Range)
    
    Dim L As Long, t As Long, W As Long, H As Long
    
    L = R.Left
    t = R.Top
    W = R.Width
    H = R.Height
    
    'msoTextOrientationHorizontal
    With ActiveSheet.Shapes
        .AddTextbox(msoTextOrientationVertical, L, t, W, H).Select
        Selection.ShapeRange.Line.Visible = msoFalse
    End With
        
End Sub

Public Sub PrintPDF()

    On Error GoTo PrintPDF_Error

    ActiveSheet.PageSetup.Zoom = False
    ActiveSheet.PageSetup.BlackAndWhite = Not tbl_Input.cb_print_color

    [input_print_area].ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=CStr([input_object_address] & "_" & [input_calculation_date]), _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True

    On Error GoTo 0
    Exit Sub

PrintPDF_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPDF of Modul mod_Drucken"

End Sub

Public Sub PrintPage()

    Dim Sh                      As Worksheet
    Dim rngPrint                As Range
    Dim s_reduce_paper_title    As String

    On Error GoTo PrintPage_Error

    s_reduce_paper_title = "Reduzieren Sie den Papierverbrauch"
    ActiveSheet.PageSetup.BlackAndWhite = Not tbl_Input.cb_print_color
    
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

Public Sub ChangeCaption(lng_message As Long)

    Select Case lng_message
        Case 0:
            Application.Caption = "Currently running"
        Case 1:
            Application.Caption = "Nicht erfolgreich"
        Case 2:
            Application.Caption = "Erfolg"
        Case Else:
            Application.Caption = "Unknown"
    End Select
End Sub

Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Application.Calculation = xlAutomatic
    
    Call ProtectPAKU2

End Sub

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlAutomatic
    
    ActiveWindow.View = xlNormalView
    Call UnProtectPAKU2

End Sub

Public Sub DeleteName(sName As String)

   On Error GoTo DeleteName_Error

    ActiveWorkbook.Names(sName).Delete
    
    Debug.Print sName & " is deleted!"
    
   On Error GoTo 0
   Exit Sub

DeleteName_Error:

    Debug.Print sName & " not present or some error"
    On Error GoTo 0
    
End Sub

Public Function RGB2HTMLColor(R As Byte, G As Byte, _
                            b As Byte) As String


'INPUT: Numeric (Base 10) Values for R, G, and B)

'RETURNS:
'A string that can be used as an HTML Color
'(i.e., "#" + the Hexadecimal equivalent)

'For VBA the RGB is reversed. R and B are revered...

    Dim HexR, HexB, HexG As Variant

    On Error GoTo ErrorHandler

    'R
    HexR = Hex(R)
    If Len(HexR) < 2 Then HexR = "0" & HexR

    'Get Green Hex
    HexG = Hex(G)
    If Len(HexG) < 2 Then HexG = "0" & HexG

    HexB = Hex(b)
    If Len(HexB) < 2 Then HexB = "0" & HexB



    RGB2HTMLColor = "#" & HexR & HexG & HexB
ErrorHandler:
End Function

Public Sub SelectAndChange()
        
    Dim current_cells_range         As Range
    
    Dim l_step_between_BA           As Long
    Dim l_counter                   As Long
    Dim col                         As Long
    Dim row                         As Long
    
    l_step_between_BA = 22
    col = Selection.Column
    row = Selection.row
    'Beware what you select, for it would stay selected! :)
    
    Set current_cells_range = Selection
    
    For l_counter = 0 To 9
        Set current_cells_range = Union(current_cells_range, ActiveSheet.Cells(row + l_step_between_BA * l_counter, col))
        
    Next l_counter
    
    current_cells_range.Select
    
End Sub

Function NamedRangeExists(strRangeName As String) As Boolean
    Dim my_range As Range
    
    On Error Resume Next
    
    Set my_range = Range(strRangeName)
    
    If Not my_range Is Nothing Then NamedRangeExists = True
    
    On Error GoTo 0
    
End Function
