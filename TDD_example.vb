Option Explicit

Public Sub TDD()
    
    Call SetToZero
    Call SetToDefault
    Call tbl_main.cmd_hoai_Click
    Call RunMe(1)
    
    Call TDD_1
    Call TDD_2
    
End Sub

Public Sub TDD_1()
    
    Call TDD_1A
    Call TDD_1B
    Call TDD_1C
    
End Sub

Public Sub TDD_2()
    
    Call TDD_2A
    Call TDD_2B
    
End Sub

Public Sub TDD_2B()

    Dim my_arr                      As Variant
    Dim specs                       As New SpecSuite
    Dim l_counter                   As Long
    Dim l_size                      As Long: l_size = 4

    Dim l_row                       As Long
    Dim l_col                       As Long

    On Error Resume Next
    Call OnStart

    my_arr = arr_fill_predefined_test_2B_rng_C1F42

    For l_counter = 0 To UBound(my_arr) - 1 Step 1
    
        l_row = l_counter \ l_size
        l_col = l_counter Mod l_size

        specs.It("2B_01_" & l_row + 1 & "_" & l_col + 2).Expect(my_arr(l_counter + 1)).ToEqual tbl_calendar.[C1].Offset(l_row, l_col).value
        'Debug.Print tbl_calendar.[C1].Offset(l_row, l_col).Address
        'tbl_calendar.[C1].Offset(l_row, l_col).Select
        
    Next l_counter

    InlineRunner.RunSuite specs
    Call specs.TotalTests
    Call OnEnd

    On Error GoTo 0

End Sub

'---------------------------------------------------------------------------------------
' Method : MakeAllValues
' Author : v.doynov
' Date   : 03.11.2016
' Purpose: Select the range, for which you want the TDD code.
'---------------------------------------------------------------------------------------

Public Sub MakeAllValues()
    
    Dim my_cell                 As Range
    Dim l_counter               As Long
    Dim str                     As String
    
    For Each my_cell In Selection
        Call Increment(l_counter)
        str = vbTab & "my_arr(" & l_counter & ")= "
        
        If Len(my_cell) > 0 Then
            If IsDate(my_cell) Then
                str = str & "CDate(""" & my_cell & """)"
            Else
                If Not IsNumeric(my_cell) Then
                    str = str & """" & my_cell & """"
                Else
                    str = str & change_commas(my_cell.value)
                End If
            End If
        Else
            str = str & 0
        End If
        
        Debug.Print str
    Next my_cell
    
End Sub

Public Sub TDD_2A()

    Dim my_arr                      As Variant
    Dim specs                       As New SpecSuite
    Dim l_counter                   As Long

    On Error Resume Next
    Call OnStart

    'Col F - Honorar
    my_arr = arr_fill_predefined_test_2A_colF
    For l_counter = 1 To UBound(my_arr) Step 1
        specs.It("2A_01F_" & l_counter).Expect(my_arr(l_counter)).ToEqual tbl_calendar.[F1].Offset(l_counter - 1).value
    Next l_counter
    
    'Col I - Mar 15
    my_arr = arr_fill_predefined_test_2A_colI
    For l_counter = 1 To UBound(my_arr) Step 1
        specs.It("2A_02I_" & l_counter).Expect(my_arr(l_counter)).ToEqual tbl_calendar.[I1].Offset(l_counter - 1).value
    Next l_counter

    'Col M - Aug 15
    my_arr = arr_fill_predefined_test_2A_colM
    Call Increment(l_counter)
    For l_counter = 1 To UBound(my_arr) Step 1
        specs.It("2A_03M_" & l_counter).Expect(my_arr(l_counter)).ToEqual tbl_calendar.[M1].Offset(l_counter - 1).value
    Next l_counter

    'Col BK - Oct 19
    my_arr = arr_fill_predefined_test_2A_colBK
    For l_counter = 1 To UBound(my_arr) Step 1
        specs.It("2A_04BK_" & l_counter).Expect(my_arr(l_counter)).ToEqual tbl_calendar.[BK1].Offset(l_counter - 1).value
    Next l_counter

    'Col AL - Sep 17
    my_arr = arr_fill_predefined_test_2A_colAL
    For l_counter = 1 To UBound(my_arr) Step 1
        specs.It("2A_05AL_" & l_counter).Expect(my_arr(l_counter)).ToEqual tbl_calendar.[AL1].Offset(l_counter - 1).value
    Next l_counter

    InlineRunner.RunSuite specs
    Call specs.TotalTests
    Call OnEnd

    On Error GoTo 0

End Sub

Public Sub MakeValues()

    Dim my_cell         As Range
    Dim str             As String
    Dim l_counter       As Long
    
    For Each my_cell In Selection
        Call Increment(l_counter)
        str = "my_arr(" & l_counter & ")= "
        
        If Len(my_cell) > 0 Then
            str = str & change_commas(my_cell.value)
        Else
            str = str & 0
        End If
        
        Debug.Print str
        
    Next my_cell

End Sub

Public Sub SetToZero()

    Dim arr_dates(12)               As Date
    Dim arr_values(16)              As Double
        
    Call OnStart
        
    tbl_main.tb_show_hide_further = True
    
    tbl_main.cmb_ba = 2
    tbl_main.cmb_land = "Deutschland"
    
    tbl_main.chb_zweimal = True
    tbl_main.chb_jump = False
    tbl_main.chb_insti = False
    
    'Set dates
    tbl_main.[m_buying_date] = ""
    tbl_main.[m_end_date] = ""
    
    tbl_main.[e2] = ""
    tbl_main.[e3] = ""
        
    tbl_main.[f2] = ""
    tbl_main.[f3] = ""
        
    tbl_main.[g2] = ""
    tbl_main.[g3] = ""
    
    tbl_main.[h2] = ""
    tbl_main.[h3] = ""
    
    tbl_main.[k2] = ""
    tbl_main.[l2] = ""
    
    'Set values
    
    tbl_main.[i2] = ""
    tbl_main.[i3] = ""
    tbl_main.[j2] = ""
    tbl_main.[j3] = ""
    
    tbl_main.[e18] = ""
    tbl_main.[e19] = ""
    tbl_main.[s54] = ""
    tbl_main.[s55] = ""
    tbl_main.[t54] = ""
    tbl_main.[t55] = ""
    tbl_main.[u54] = ""
    tbl_main.[u55] = ""
    tbl_main.[v54] = ""
    tbl_main.[v55] = ""
    tbl_main.[i92] = ""
    tbl_main.[i93] = ""
    
    Call OnEnd
    'Call HOAI calculation
    
    On Error GoTo 0
    Exit Sub
    
End Sub


Public Sub SetToDefault()

    If [set_in_production] Then On Error GoTo SetToDefault_Error
    
    Dim arr_dates(12)               As Date
    Dim arr_values(16)              As Double
    
    Call OnStart
    tbl_main.tb_show_hide_further = True
    
    tbl_main.cmb_ba = 2
    tbl_main.cmb_land = "Deutschland"
    
    tbl_main.chb_zweimal = True
    tbl_main.chb_jump = False
    tbl_main.chb_insti = False
    
    'Set dates
    arr_dates(1) = "01.03.2015"
    arr_dates(2) = "01.10.2019"
    arr_dates(3) = "01.12.2016"
    arr_dates(4) = "01.12.2016"
    arr_dates(5) = "01.06.2018"
    arr_dates(6) = "01.07.2018"
    arr_dates(7) = "01.08.2018"
    arr_dates(8) = "01.10.2018"
    arr_dates(9) = "01.09.2017"
    arr_dates(10) = "01.05.2017"
    arr_dates(11) = "01.01.2016"
    arr_dates(12) = "01.07.2015"
    
    tbl_main.[main_objektname] = "Bagelstrasse Duesseldorf"
    tbl_main.[m_buying_date] = arr_dates(1)
    tbl_main.[m_end_date] = arr_dates(2)
    
    tbl_main.[e2] = arr_dates(3)
    tbl_main.[e3] = arr_dates(4)
        
    tbl_main.[f2] = arr_dates(5)
    tbl_main.[f3] = arr_dates(6)
        
    tbl_main.[g2] = arr_dates(7)
    tbl_main.[g3] = arr_dates(8)
    
    tbl_main.[h2] = arr_dates(9)
    tbl_main.[h3] = arr_dates(10)
    
    tbl_main.[k2] = arr_dates(11)
    tbl_main.[l2] = arr_dates(12)
    
    'Set values
    arr_values(1) = 3417
    arr_values(2) = 3644
    arr_values(3) = 404
    arr_values(4) = 404
    arr_values(5) = 1234567
    arr_values(6) = 12345678
    arr_values(7) = 123456
    arr_values(8) = 100000
    arr_values(9) = 250000
    arr_values(10) = 270000
    arr_values(11) = 350000
    arr_values(12) = 450000
    arr_values(13) = 300000
    arr_values(14) = 350000
    arr_values(15) = 150000
    arr_values(16) = 160000
    
    tbl_main.[i2] = arr_values(1)
    tbl_main.[i3] = arr_values(2)
    tbl_main.[j2] = arr_values(3)
    tbl_main.[j3] = arr_values(4)
    
    tbl_main.[e18] = arr_values(5)
    tbl_main.[e19] = arr_values(6)
    tbl_main.[s54] = arr_values(7)
    tbl_main.[s55] = arr_values(8)
    tbl_main.[t54] = arr_values(9)
    tbl_main.[t55] = arr_values(10)
    tbl_main.[u54] = arr_values(11)
    tbl_main.[u55] = arr_values(12)
    tbl_main.[v54] = arr_values(13)
    tbl_main.[v55] = arr_values(14)
    tbl_main.[i92] = arr_values(15)
    tbl_main.[i93] = arr_values(16)
    
    Call OnEnd
    'Call HOAI calculation

    On Error GoTo 0
    Exit Sub

SetToDefault_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetToDefault of Sub mod_TDD"

End Sub

Public Sub HowToList()

    Dim obj_list As cls_vbaList

    Set obj_list = New cls_vbaList

    obj_list.Add (30)
    obj_list.Add (3)
    obj_list.Add (355)
    obj_list.Add (5)
    obj_list.Add (1)
    obj_list.Add (40)

    Debug.Print obj_list.Contains(30)
    Debug.Print obj_list.Exists(30)
    Debug.Print obj_list.Items(0)
    
    obj_list.Sort
    
    Debug.Print obj_list.Items(0)
    Debug.Print obj_list.Find(3)
    Debug.Print obj_list.Find(30)
    Debug.Print obj_list.LastIndexOf(355)

    Set obj_list = Nothing

End Sub

Public Sub TDD_1C()

    On Error Resume Next
    
    Dim specs                       As New SpecSuite
    
    Dim obj_total_test              As New cls_Total
    Dim obj_total_cal_test          As New cls_TotalCalendar
    Dim var_list                    As New cls_vbaList
    
    Call OnStart
    
    specs.It("C001").Expect(obj_total_test.LeftSideCols).ToEqual 7
    specs.It("C002").Expect(obj_total_test.BA_Number).ToEqual CLng(tbl_main.[cmb_ba].value)
    specs.It("C003").Expect(obj_total_test.B_Insti).ToEqual CBool(tbl_main.[chb_insti])
    
    tbl_main.[chb_insti] = True
    specs.It("C004").Expect(obj_total_test.MarkCost1).ToEqual CStr([set_total_mark_2])
    specs.It("C005").Expect(obj_total_test.MarkCost2).ToEqual CStr([set_total_mark_4])
    specs.It("C006").Expect(obj_total_test.MarkCost3).ToEqual CStr([set_total_mark_6])
    
    tbl_main.[chb_insti] = False
    specs.It("C007").Expect(obj_total_test.MarkCost1).ToEqual CStr([set_total_mark_1])
    specs.It("C008").Expect(obj_total_test.MarkCost2).ToEqual CStr([set_total_mark_3])
    specs.It("C009").Expect(obj_total_test.MarkCost3).ToEqual CStr([set_total_mark_5])
    
    specs.It("C010").Expect(obj_total_test.MarkCost1).ToNotEqual CStr([set_total_mark_2])
    specs.It("C011").Expect(obj_total_test.MarkCost2).ToNotEqual CStr([set_total_mark_4])
    specs.It("C012").Expect(obj_total_test.MarkCost3).ToNotEqual CStr([set_total_mark_6])
    
    specs.It("C013").Expect(obj_total_test.CurrentLine).ToNotEqual 2
    specs.It("C014").Expect(obj_total_test.CurrentLine).ToEqual 0
    Call obj_total_test.IncrementCurrentLine
    specs.It("C015").Expect(obj_total_test.CurrentLine).ToEqual 1
    Call obj_total_test.IncrementCurrentLine
    specs.It("C016").Expect(obj_total_test.CurrentLine).ToEqual 2
    obj_total_test.CurrentLine = 12
    specs.It("C017").Expect(obj_total_test.CurrentLine).ToEqual 12
    
    specs.It("C018").Expect(obj_total_test.MarkCostTitle).ToEqual CStr([set_total_mark_title])
    specs.It("C019").Expect(obj_total_test.CurrentLine).ToNotEqual CStr([set_total_mark_title] & "1")
    
    specs.It("C020").Expect(obj_total_test.LastRow).ToEqual last_row(tbl_totals.Name)
    specs.It("C021").Expect(obj_total_test.LastRow).ToNotEqual 9999
    
    var_list.Add 5
    var_list.Add 10
    var_list.Add 15
    var_list.Add 20
    
    Set obj_total_cal_test.PlanerkostenSrc = var_list
    
    specs.It("C022").Expect(obj_total_cal_test.PlanerkostenSrc.Items(0)).ToEqual 5
    specs.It("C023").Expect(obj_total_cal_test.PlanerkostenSrc.Items(1)).ToEqual 10
    specs.It("C024").Expect(obj_total_cal_test.PlanerkostenSrc.Items(2)).ToEqual 15
    specs.It("C025").Expect(obj_total_cal_test.PlanerkostenSrc.Items(3)).ToEqual 20
    
    specs.It("C026").Expect(obj_total_cal_test.PlanerkostenSrc.Items(0)).ToNotEqual 5 + 1
    specs.It("C027").Expect(obj_total_cal_test.PlanerkostenSrc.Items(1)).ToNotEqual 10 + 1
    specs.It("C028").Expect(obj_total_cal_test.PlanerkostenSrc.Items(2)).ToNotEqual 15 + 1
    
    InlineRunner.RunSuite specs
    
    Call specs.TotalTests
    Call OnEnd
    
    On Error GoTo 0
    
End Sub


Public Sub TDD_1B()
    
    On Error Resume Next

    Dim specs               As New SpecSuite
    
    Call OnStart
    
    specs.It("B001").Expect([set_in_production]).ToEqual True
    specs.It("B002").Expect([set_in_production]).ToNotEqual False
    
    InlineRunner.RunSuite specs
    
    Call specs.TotalTests
    
    Call OnEnd
    
    On Error GoTo 0
    
End Sub

Public Sub TDD_1A()

    On Error Resume Next

    Dim specs               As New SpecSuite
    Dim obj_calendar        As New cls_Calendar
    Dim obj_dat             As New cls_Dates
    Dim obj_sav             As New cls_Saver
    Dim obj_input_dates     As New cls_InputDates
    Dim obj_test_land       As New cls_Land
    
    Dim l_value             As Long
    Dim d_value             As Date
    Dim str_initial         As String
    
    Call OnStart
    Set obj_con = New cls_Const
    
    specs.It("A001").Expect(obj_calendar.UPPER_ROW).ToEqual 4
    specs.It("A002").Expect(obj_calendar.ROWS_TAKEN).ToEqual 3

    obj_calendar.current_row = 111
    specs.It("A003").Expect(obj_calendar.current_row).ToEqual 111

    obj_calendar.IncrementRow
    obj_calendar.IncrementRow
    specs.It("A004").Expect(obj_calendar.current_row).ToNotEqual 111
    
    obj_calendar.IncrementRow
    specs.It("A005").Expect(obj_calendar.current_row).ToEqual 114
    
    obj_calendar.AddToPercentageLines (10)
    obj_calendar.AddToPercentageLines (15)
    obj_calendar.AddToPercentageLines (20)
    obj_calendar.AddToPercentageLines (25)

    specs.It("A006").Expect(obj_calendar.percentage_lines(1)).ToEqual 10
    specs.It("A007").Expect(obj_calendar.percentage_lines(2)).ToEqual 15
    specs.It("A008").Expect(obj_calendar.percentage_lines(3)).ToNotEqual 20 + 1
    specs.It("A009").Expect(obj_calendar.percentage_lines(4)).ToEqual 25

    obj_calendar.AddToLines (100)
    obj_calendar.AddToLines (200)
    obj_calendar.AddToLines (300)
    obj_calendar.AddToLines (400)

    specs.It("A010").Expect(obj_calendar.lines(1)).ToEqual 100
    specs.It("A011").Expect(obj_calendar.lines(2)).ToEqual 200
    specs.It("A012").Expect(obj_calendar.lines(3)).ToNotEqual 300 + 1
    specs.It("A013").Expect(obj_calendar.lines(3)).ToEqual 300
    specs.It("A014").Expect(obj_calendar.lines(4)).ToEqual 400

    specs.It("A015").Expect(obj_calendar.lines(4)).ToEqual 400

    obj_calendar.last_col = 400
    specs.It("A016").Expect(obj_calendar.length_of_calendar).ToEqual 400 - obj_con.COLUMNS_TAKEN

    Dim str_variable As String: str_variable = "BA LP"
    specs.It("A017").Expect(obj_con.BA_NAME & obj_con.SPACE & obj_con.LP_NAME).ToEqual (str_variable)

    str_variable = "BA L P"
    specs.It("A018").Expect(obj_con.BA_NAME & obj_con.SPACE & obj_con.LP_NAME).ToNotEqual (str_variable)

    specs.It("A019").Expect(generate_honorare_gebaude(100000, 3, True)).ToEqual 15005
    specs.It("A020").Expect(generate_honorare_gebaude(100000, 3, False)).ToEqual 16859
    specs.It("A021").Expect(generate_honorare_gebaude(100000, 3, True)).ToNotEqual 15005 + 10
    specs.It("A022").Expect(generate_honorare_gebaude(100000, 3, False)).ToNotEqual 16859 + 10
    
    specs.It("A023").Expect(generate_honorare_hlse(100000, 2, True)).ToEqual 27150
    specs.It("A024").Expect(generate_honorare_hlse(100000, 2, False)).ToEqual 29511
    specs.It("A025").Expect(generate_honorare_hlse(5000, 2, True)).ToEqual 2547
    specs.It("A026").Expect(generate_honorare_hlse(5000, 2, False)).ToEqual 2768.5
    specs.It("A027").Expect(generate_honorare_hlse(4000000, 2, True)).ToEqual 492410
    specs.It("A028").Expect(generate_honorare_hlse(4000000, 2, False)).ToEqual 535228
    specs.It("A029").Expect(generate_honorare_hlse(4000000, 2, True)).ToNotEqual 492410 - 10
    specs.It("A030").Expect(generate_honorare_hlse(4000000, 2, False)).ToNotEqual 535228 - 10
    
    specs.It("A031").Expect(generate_honorare_aussenanlagen(20000, 3, True)).ToEqual 5229
    specs.It("A032").Expect(generate_honorare_aussenanlagen(20000, 3, False)).ToEqual 5875
    specs.It("A033").Expect(generate_honorare_aussenanlagen(75000, 3, True)).ToEqual 16116
    specs.It("A034").Expect(generate_honorare_aussenanlagen(75000, 3, False)).ToEqual 18108
    specs.It("A035").Expect(generate_honorare_aussenanlagen(1500000, 3, True)).ToEqual 201261
    specs.It("A036").Expect(generate_honorare_aussenanlagen(1500000, 3, False)).ToEqual 226136
    specs.It("A037").Expect(generate_honorare_aussenanlagen(1500000, 3, True)).ToNotEqual 201261 + 10
    specs.It("A038").Expect(generate_honorare_aussenanlagen(1500000, 3, False)).ToNotEqual 226132 + 10

    specs.It("A039").Expect(generate_honorare_tragwerksplannung(10000, 3, True)).ToEqual 2064
    specs.It("A040").Expect(generate_honorare_tragwerksplannung(10000, 3, False)).ToEqual 2319.5
    specs.It("A041").Expect(generate_honorare_tragwerksplannung(123456, 3, True)).ToEqual 14863.1
    specs.It("A042").Expect(generate_honorare_tragwerksplannung(123456, 3, False)).ToEqual 16700.24
    specs.It("A043").Expect(generate_honorare_tragwerksplannung(15000000, 3, True)).ToEqual 642943
    specs.It("A044").Expect(generate_honorare_tragwerksplannung(15000000, 3, False)).ToEqual 722408
    specs.It("A045").Expect(generate_honorare_tragwerksplannung(15000000, 3, True)).ToNotEqual 642943 + 1
    specs.It("A046").Expect(generate_honorare_tragwerksplannung(15000000, 3, False)).ToNotEqual 722408 + 1
    
    specs.It("A047").Expect(generate_honorar_brandschutz(969)).ToEqual 8994.56
    specs.It("A048").Expect(generate_honorar_brandschutz(2322)).ToEqual 13652.83
    specs.It("A049").Expect(generate_honorar_brandschutz(12345.67)).ToEqual 33544.66
    specs.It("A050").Expect(generate_honorar_brandschutz(25900.18)).ToEqual 51136.09
    
    '   b_show_msgbox is an optional value, set to true initially.
    '   The idea is to be false for the tests, thus it does not show a msgbox
    specs.It("A051").Expect(generate_honorare_tragwerksplannung(10000 - 1, 3, True, b_show_msgbox:=False)).ToEqual -1
    specs.It("A052").Expect(generate_honorare_tragwerksplannung(15000000 + 1, 3, True, b_show_msgbox:=False)).ToEqual -10

    obj_calendar.last_col = 50
    specs.It("A053").Expect(obj_calendar.last_col).ToEqual 50
    
    Set obj_dat = New cls_Dates
    Call obj_dat.AddEingabeDate("04.02.1999")
    Call obj_dat.AddEingabeDate("04.02.1998")
    Call obj_dat.AddEingabeDate("04.02.1995")
    
    specs.It("A054").Expect(obj_dat.eingabe_date(1)).ToEqual CDate("04.02.1999")
    specs.It("A055").Expect(obj_dat.eingabe_date(2)).ToEqual CDate("04.02.1998")
    specs.It("A056").Expect(obj_dat.eingabe_date(3)).ToEqual CDate("04.02.1995")
    specs.It("A057").Expect(obj_dat.eingabe_date(2)).ToEqual CDate("04.02.1998")
    specs.It("A058").Expect(obj_dat.eingabe_date(3)).ToNotEqual CDate("05.02.1999")

    obj_calendar.last_col = obj_calendar.last_col + obj_calendar.last_col
    specs.It("A059").Expect(obj_calendar.last_col).ToEqual 100

    l_value = tbl_main.cmb_ba.value
    specs.It("A060").Expect(obj_calendar.ba).ToEqual l_value
    specs.It("A061").Expect(obj_calendar.ba).ToNotEqual l_value + 1

    d_value = DateSerial(tbl_main.cmb_year, tbl_main.cmb_month, 1)
    specs.It("A062").Expect(obj_calendar.fixed_date).ToEqual d_value
    specs.It("A063").Expect(obj_calendar.fixed_date).ToNotEqual d_value + 1

    d_value = DateDiff("m", [m_start_date], [m_end_date])
    specs.It("A064").Expect(obj_calendar.calendar_size_original).ToEqual d_value
    specs.It("A065").Expect(obj_calendar.calendar_size_original).ToNotEqual d_value + 1

    d_value = DateDiff("m", [m_start_date], [main_bau_range_changes_2])
    specs.It("A066").Expect(obj_calendar.calendar_size_changed).ToEqual d_value
    specs.It("A067").Expect(obj_calendar.calendar_size_changed).ToNotEqual d_value + 1
    
    Set obj_sav = New cls_Saver
    obj_sav.AddRate7 ("12.12.2012")
    obj_sav.AddRate7 ("12.12.2013")
    obj_sav.AddRate7 ("12.12.2014")
    
    specs.It("A068").Expect(obj_sav.Rate7MF(1)).ToEqual CDate("12.12.2012")
    specs.It("A069").Expect(obj_sav.Rate7MF(2)).ToEqual CDate("12.12.2013")
    specs.It("A070").Expect(obj_sav.Rate7MF(3)).ToEqual CDate("12.12.2014")
    specs.It("A071").Expect(obj_sav.Rate7MF(1)).ToNotEqual CDate("12.12.2015")
    
    obj_sav.AddRate6 ("12.5.2012")
    obj_sav.AddRate6 ("12.6.2013")
    obj_sav.AddRate6 ("12.7.2014")

    specs.It("A072").Expect(obj_sav.Rate6BZ(1)).ToEqual CDate("12.5.2012")
    specs.It("A073").Expect(obj_sav.Rate6BZ(2)).ToEqual CDate("12.6.2013")
    specs.It("A074").Expect(obj_sav.Rate6BZ(3)).ToEqual CDate("12.7.2014")
    specs.It("A075").Expect(obj_sav.Rate6BZ(1)).ToNotEqual CDate("12.12.2015")
    
    obj_sav.AddBB ("12.1.2012")
    obj_sav.AddBB ("12.2.2013")
    obj_sav.AddBB ("12.3.2014")
    
    specs.It("A076").Expect(obj_sav.BB(1)).ToEqual CDate("12.1.2012")
    specs.It("A077").Expect(obj_sav.BB(2)).ToEqual CDate("12.2.2013")
    specs.It("A078").Expect(obj_sav.BB(3)).ToEqual CDate("12.3.2014")
    specs.It("A079").Expect(obj_sav.BB(1)).ToNotEqual CDate("12.4.2015")
    
    obj_sav.AddEndeRb ("1.5.2012")
    obj_sav.AddEndeRb ("2.5.2012")
    obj_sav.AddEndeRb ("3.5.2012")
    
    specs.It("A080").Expect(obj_sav.EndeRb(1)).ToEqual CDate("1.5.2012")
    specs.It("A081").Expect(obj_sav.EndeRb(2)).ToEqual CDate("2.5.2012")
    specs.It("A082").Expect(obj_sav.EndeRb(3)).ToEqual CDate("3.5.2012")
    specs.It("A083").Expect(obj_sav.EndeRb(1)).ToNotEqual CDate("2.5.2012")
    
    obj_sav.Baueingabe = "6.10.2020"
    specs.It("A084").Expect(obj_sav.Baueingabe).ToEqual CDate("6.10.2020")
    obj_sav.Baueingabe = "6.10.2021"
    specs.It("A085").Expect(obj_sav.Baueingabe).ToEqual CDate("6.10.2021")
    obj_sav.Baueingabe = "6.10.2022"
    specs.It("A086").Expect(obj_sav.Baueingabe).ToEqual CDate("6.10.2022")
    specs.It("A087").Expect(obj_sav.Baueingabe).ToNotEqual CDate("6.10.2023")
    
    obj_sav.Baugenehmigung = "6.11.2020"
    specs.It("A088").Expect(obj_sav.Baugenehmigung).ToEqual CDate("6.11.2020")
    obj_sav.Baugenehmigung = "6.11.2021"
    specs.It("A089").Expect(obj_sav.Baugenehmigung).ToEqual CDate("6.11.2021")
    obj_sav.Baugenehmigung = "6.11.2022"
    specs.It("A090").Expect(obj_sav.Baugenehmigung).ToEqual CDate("6.11.2022")
    specs.It("A091").Expect(obj_sav.Baugenehmigung).ToNotEqual CDate("6.11.2023")
    
    obj_sav.LetzterTag = "12.12.1960"
    specs.It("A092").Expect(obj_sav.LetzterTag).ToEqual CDate("12.12.1960")
    obj_sav.LetzterTag = "12.12.1961"
    specs.It("A093").Expect(obj_sav.LetzterTag).ToEqual CDate("12.12.1961")
    obj_sav.LetzterTag = "12.12.1962"
    specs.It("A094").Expect(obj_sav.LetzterTag).ToEqual CDate("12.12.1962")
    specs.It("A095").Expect(obj_sav.LetzterTag).ToNotEqual CDate("12.12.1960")

    obj_sav.Changes = "vit"
    specs.It("A096").Expect(obj_sav.Changes).ToEqual "vit"
    obj_sav.Changes = "osh"
    specs.It("A097").Expect(obj_sav.Changes).ToNotEqual "vit"
    specs.It("A098").Expect(obj_sav.Changes).ToEqual "vit" & vbCrLf & "osh"

    obj_sav.AddChangeCell ("Pesho beshe tuk")
    obj_sav.AddChangeCell ("Gosho beshe tuk")
    obj_sav.AddChangeCell ("Atanas beshe tuk")
    obj_sav.AddChangeCell ("I az byah tuk")

    specs.It("A099").Expect(obj_sav.ChangeCell(1)).ToEqual "Pesho beshe tuk"
    specs.It("A100").Expect(obj_sav.ChangeCell(2)).ToEqual "Gosho beshe tuk"
    specs.It("A101").Expect(obj_sav.ChangeCell(3)).ToEqual "Atanas beshe tuk"
    specs.It("A102").Expect(obj_sav.ChangeCell(3)).ToNotEqual "Gosho beshe tuk"
    
    specs.It("A103").Expect(obj_sav.ChangesTotal).ToEqual 4
    specs.It("A104").Expect(obj_sav.ChangesTotal).ToNotEqual 5
    obj_sav.AddChangeCell ("I az byah tuk2")
    specs.It("A105").Expect(obj_sav.ChangesTotal).ToEqual 5
    
    specs.It("A106").Expect(obj_sav.Changes).ToEqual "vit" & vbCrLf & "osh"
    obj_sav.EraseChanges
    specs.It("A107").Expect(obj_sav.Changes).ToEqual ""
    obj_sav.Changes = "vi"
    obj_sav.Changes = "to"
    specs.It("A108").Expect(obj_sav.Changes).ToEqual "vi" & vbCrLf & "to"
        
    specs.It("A109").Expect(obj_con.FORMULA_CALCULATIONS(10, 5, 3, True)).ToEqual "=RC[-1]+(((RC6-RC10)*0.9)/5)"
    specs.It("A110").Expect(obj_con.FORMULA_CALCULATIONS(10, 5, 3)).ToEqual "=RC[-1]+((RC6-RC10)/5)"
    
    specs.It("A111").Expect(obj_con.FORMULA_CALCULATIONS(10, 5, 0, True)).ToEqual "=RC[-1]+(RC6*0.9/5)"
    specs.It("A112").Expect(obj_con.FORMULA_CALCULATIONS(10, 5, 0)).ToEqual "=RC[-1]+(RC6/5)"
    
    obj_input_dates.AddRate1_Date ("01.01.2013")
    obj_input_dates.AddRate1_Date ("02.01.2013")
    obj_input_dates.AddRate1_Date ("03.01.2013")
    specs.It("A113").Expect(obj_input_dates.rate1_date(3)).ToEqual CDate("03.01.2013")
    specs.It("A114").Expect(obj_input_dates.rate1_date(1)).ToNotEqual CDate("02.01.2013")
    specs.It("A115").Expect(obj_input_dates.rate1_date(2)).ToEqual CDate("02.01.2013")
        
    obj_input_dates.AddRate2_Date ("01.01.2011")
    obj_input_dates.AddRate2_Date ("01.01.2012")
    obj_input_dates.AddRate2_Date ("01.01.2013")
    obj_input_dates.AddRate2_Date ("01.01.2014")
    obj_input_dates.AddRate2_Date ("01.01.2015")
    obj_input_dates.AddRate2_Date ("01.01.2016")
    specs.It("A116").Expect(obj_input_dates.rate2_date(1)).ToEqual CDate("01.01.2011")
    specs.It("A117").Expect(obj_input_dates.rate2_date(4)).ToEqual CDate("01.01.2014")
    specs.It("A118").Expect(obj_input_dates.rate2_date(5)).ToNotEqual CDate("01.01.2014")
    
    obj_input_dates.AddRate6_Date ("01.01.2020")
    obj_input_dates.AddRate6_Date ("01.01.2021")
    obj_input_dates.AddRate6_Date ("01.01.2022")
    specs.It("A119").Expect(obj_input_dates.rate6_date(3)).ToEqual CDate("01.01.2022")
    specs.It("A120").Expect(obj_input_dates.rate6_date(1)).ToEqual CDate("01.01.2020")
    specs.It("A121").Expect(obj_input_dates.rate6_date(2)).ToNotEqual CDate("01.01.2020")
    
    obj_input_dates.AddRate7_Date ("01.01.2013")
    obj_input_dates.AddRate7_Date ("02.01.2013")
    obj_input_dates.AddRate7_Date ("03.01.2013")
    obj_input_dates.AddRate7_Date ("04.01.2013")
    obj_input_dates.AddRate7_Date ("05.01.2013")
    obj_input_dates.AddRate7_Date ("06.01.2013")
    specs.It("A122").Expect(obj_input_dates.rate7_date(6)).ToEqual CDate("06.01.2013")
    specs.It("A123").Expect(obj_input_dates.rate7_date(5)).ToEqual CDate("05.01.2013")
    specs.It("A124").Expect(obj_input_dates.rate7_date(6)).ToEqual CDate("06.01.2013")
    specs.It("A125").Expect(obj_input_dates.rate7_date(5)).ToEqual CDate("05.01.2013")
    specs.It("A126").Expect(obj_input_dates.rate7_date(1)).ToEqual CDate("01.01.2013")
    specs.It("A127").Expect(obj_input_dates.rate7_date(2)).ToNotEqual CDate("01.01.2013")
    
    obj_input_dates.Ankaufsdatum = CDate("07.08.2011")
    specs.It("A128").Expect(obj_input_dates.Ankaufsdatum).ToEqual CDate("07.08.2011")
    obj_input_dates.Ankaufsdatum = CDate("07.08.2012")
    specs.It("A129").Expect(obj_input_dates.Ankaufsdatum).ToEqual CDate("07.08.2012")
    specs.It("A130").Expect(obj_input_dates.Ankaufsdatum).ToNotEqual CDate("07.08.2013")
    obj_input_dates.Ankaufsdatum = CDate("07.08.2013")
    specs.It("A131").Expect(obj_input_dates.Ankaufsdatum).ToEqual CDate("07.08.2013")

    obj_input_dates.Baueingabe = CDate("07.08.2021")
    specs.It("A132").Expect(obj_input_dates.Baueingabe).ToEqual CDate("07.08.2021")
    obj_input_dates.Baueingabe = CDate("07.08.2022")
    specs.It("A133").Expect(obj_input_dates.Baueingabe).ToEqual CDate("07.08.2022")
    obj_input_dates.Baueingabe = CDate("07.08.2023")
    specs.It("A134").Expect(obj_input_dates.Baueingabe).ToNotEqual CDate("07.08.2022")
    
    obj_input_dates.Baugenehmigung = CDate("01.08.2011")
    specs.It("A135").Expect(obj_input_dates.Baugenehmigung).ToEqual CDate("01.08.2011")
    obj_input_dates.Baugenehmigung = CDate("02.08.2012")
    specs.It("A136").Expect(obj_input_dates.Baugenehmigung).ToEqual CDate("02.08.2012")
    obj_input_dates.Baugenehmigung = CDate("03.08.2013")
    specs.It("A137").Expect(obj_input_dates.Baugenehmigung).ToEqual CDate("03.08.2013")
    specs.It("A138").Expect(obj_input_dates.Baugenehmigung).ToNotEqual CDate("02.08.2013")
    
    str_initial = tbl_main.cmb_land
    tbl_main.cmb_land = [set_nameGermany]
    specs.It("A139").Expect([set_vat_used].Text).ToEqual ([set_vatGermany].Text)
    specs.It("A140").Expect(obj_test_land.str_get_land).ToEqual ([set_nameGermany].Text)
    specs.It("A141").Expect(obj_test_land.str_get_short_name).ToEqual ([set_shortNameGermany].Text)
    specs.It("A142").Expect(obj_test_land.str_get_short_name).ToNotEqual ([set_shortNameAustria].Text)
    
    tbl_main.cmb_land = [set_nameAustria]

    specs.It("A143").Expect([set_vat_used].Text).ToEqual ([set_vatAustria].Text)
    specs.It("A144").Expect(obj_test_land.str_get_land).ToEqual ([set_nameAustria].Text)
    specs.It("A145").Expect(obj_test_land.str_get_short_name).ToEqual ([set_shortNameAustria].Text)
    specs.It("A146").Expect(obj_test_land.str_get_short_name).ToNotEqual ([set_shortNameGermany].Text)
    
    tbl_main.cmb_land = [set_nameGermany]
    
    InlineRunner.RunSuite specs
    Call specs.TotalTests
    
    Call OnEnd
    
    Set specs = Nothing
    Set obj_calendar = Nothing
    Set obj_con = Nothing
    Set obj_dat = Nothing
    Set obj_sav = Nothing
    Set obj_input_dates = Nothing
    Set obj_test_land = Nothing
    
    On Error GoTo 0
    
End Sub
