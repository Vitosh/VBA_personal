Public Sub remove_space_in_string()

    Dim r_range As Range
        
    For Each r_range In Selection
        r_range = Trim(r_range)
        r_range = Replace(r_range, vbTab, "")
        r_range = Replace(r_range, " ", "")
        r_range = Replace(r_range, Chr(160), "")
    Next r_range

End Sub

Public Sub FreezeTopRow()

    Dim ws          As Worksheet
    
    Application.ScreenUpdating = False
    Set ws = Worksheets("calendar")
    
    Application.Goto ws.Range("h2")
    ActiveWindow.FreezePanes = True
    
    Set ws = Nothing

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

Public Sub pls(Optional b_unhide As Boolean = False)
    If b_value_in_array(Environ("username"), S_ADMINS, True) Then
        tbl_main.Unprotect Password:=s_co
        If b_unhide Then Call UnhideAll
        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
        Debug.Print "ok :)"
    Else
        MsgBox Environ("username") & " you are not allowed to do this. Speak with Vitosh.", vbInformation, [set_planerkostenberechnung]
    End If
End Sub

Public Sub LockMe()

    tbl_main.Protect Password:=s_co
    Debug.Print "locked"
    
End Sub


Public Sub HideNeeded()
    
    Dim var_Sheet                   As Variant
    
    Dim arr_visible_sheets          As Variant
    Dim arr_hidden_sheets           As Variant
    
    Call OnStart
     
    arr_visible_sheets = Array(tbl_main, tbl_calendar)
    arr_hidden_sheets = Array(tbl_hon_aus, tbl_hon_geb, tbl_hon_hlse, tbl_hon_tra, tbl_hono_bs, tbl_hono_ps, tbl_public, tbl_settings)
    
    For Each var_Sheet In arr_visible_sheets
        var_Sheet.Visible = xlSheetVisible
    Next var_Sheet
    
    For Each var_Sheet In arr_hidden_sheets
        var_Sheet.Visible = xlSheetVeryHidden
    Next var_Sheet
   
    Call OnEnd
    
End Sub


Public Sub UnhideAll()
        
    Dim Sheet As Worksheet
    
    For Each Sheet In ThisWorkbook.Worksheets
       ' If Sheet.Visible = Not xlSheetVisible Then Sheet.Visible = xlSheetVisible
       Sheet.Visible = xlSheetVisible
    Next Sheet
    
End Sub

Public Function calculate_range(from_row As Long, to_row As Long, l_column As Long, _
                                Optional s_sheet_name As String = "calendar") As Double

    Dim ws              As Worksheet
    Dim l_counter       As Long
    Dim d_result        As Double
    
    Set ws = ThisWorkbook.Worksheets(s_sheet_name)
    
    For l_counter = from_row To to_row
        Call Increment(d_result, ws.Cells(l_counter, l_column))
    Next l_counter

    Set ws = Nothing
    
    calculate_range = Round(d_result, 2)
    
End Function


Public Sub FixOutlook()

    tbl_calendar.Cells.EntireColumn.AutoFit
   
End Sub

Public Sub HideRange(r_range_to_hide As Range)

    Dim my_cell             As Range
    Dim l_ba_value          As Long
    
    l_ba_value = tbl_main.cmb_ba.value + r_range_to_hide.Row - 1

    For Each my_cell In r_range_to_hide
        If my_cell.Row > l_ba_value Then
            my_cell.Interior.Pattern = xlGray8
            my_cell.Font.ThemeColor = xlThemeColorDark1
        Else
            my_cell.Interior.Pattern = xlAutomatic
            my_cell.Font.ColorIndex = xlAutomatic
        End If
    Next my_cell
     
    r_range_to_hide.Borders(xlEdgeTop).LineStyle = xlContinuous
    r_range_to_hide.Borders(xlEdgeLeft).LineStyle = xlContinuous
    r_range_to_hide.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r_range_to_hide.Borders(xlEdgeRight).LineStyle = xlContinuous

End Sub


Public Function add_months(ByVal my_date As Date, ByVal i_month As Integer, Optional ByVal b_use_last_date = False) As Date

    If b_use_last_date Then
        add_months = get_last_day_of_month(DateAdd("m", i_month, my_date))
    Else
        add_months = DateAdd("m", i_month, my_date)
    End If

End Function

Public Function get_last_day_of_month(my_date As Date) As Date

    get_last_day_of_month = DateSerial(Year(my_date), Month(my_date) + 1, 0)

End Function



Public Sub AddSomething(str_to_add As String, Optional c_range As Variant)
    
    Dim my_cell As Range
    
    If IsMissing(c_range) Then Set c_range = Selection
    
    For Each my_cell In c_range
        my_cell = my_cell & str_to_add
    Next my_cell
    
    Set c_range = Nothing
    
End Sub

Public Sub Meter2()

    Selection.NumberFormat = "0"" m" & Chr(179) & """"

End Sub

Public Function change_commas(ByVal myValue As Variant) As String
    
    Dim str_temp As String
    
    str_temp = CStr(myValue)
    change_commas = Replace(str_temp, ",", ".")
    
End Function

Public Sub Increment(ByRef value_to_increment, Optional l_plus As Double = 1) 'optional value type changed to double
    
    value_to_increment = value_to_increment + l_plus
    
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

Public Function RGB2HTMLColor(B As Byte, G As Byte, r As Byte) As String

    Dim HexR As Variant, HexB As Variant, HexG As Variant
    Dim sTemp As String

    On Error GoTo ErrorHandler

    'R
    HexR = Hex(r)
    If Len(HexR) < 2 Then HexR = "0" & HexR

    'Get Green Hex
    HexG = Hex(G)
    If Len(HexG) < 2 Then HexG = "0" & HexG

    HexB = Hex(B)
    If Len(HexB) < 2 Then HexB = "0" & HexB

    RGB2HTMLColor = HexR & HexG & HexB
    Debug.Print "Red and Blue are reversed ... pay attention to the input in the input"
    Exit Function
ErrorHandler:
    Debug.Print "RGB2HTMLColor was not successful"
End Function

Public Function sum_array(my_array As Variant, Optional last_values_not_to_calculate As Long = 0) As Double
    
    Dim l_counter       As Long
    
    For l_counter = LBound(my_array) To UBound(my_array) - last_values_not_to_calculate
        sum_array = sum_array + my_array(l_counter)
    Next l_counter

End Function

Public Function b_value_in_array(my_value As Variant, _
                                 my_array As Variant, _
                    Optional b_is_string As Boolean = False, _
                    Optional str_separator As String = ":") As Boolean

    Dim l_counter

    If b_is_string Then
        my_array = Split(my_array, str_separator)
    End If

    For l_counter = LBound(my_array) To UBound(my_array)
        my_array(l_counter) = CStr(my_array(l_counter))
    Next l_counter

    b_value_in_array = Not IsError(Application.Match(CStr(my_value), my_array, 0))
    
End Function


Public Sub HideSelectedSheets()
    ActiveWindow.SelectedSheets.Visible = False
End Sub
