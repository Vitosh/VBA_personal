Option Explicit

Public Sub SetGlobalVerkaufGewerbeArea(Optional b_clear_global_gewerbe = False)

    Dim my_cell             As Range
    Dim l_steps             As Long

    Call OnStart

    l_steps = 0

    If b_clear_global_gewerbe Then
        tbl_Input.[input_global_verkauf].Clear
        tbl_Input.[input_miete_range].Clear
        tbl_Input.[input_total_miete_and_kaufpreis_faktor].Clear
        tbl_Input.[input_miete_range_3_garages].Clear

        Call BlackOutRange(True)
        tbl_Input.Cells(1, 1).Select

        Call OnEnd

        Exit Sub
    End If

    Call BlackOutRange

    If tbl_Input.chb_global Then
        For Each my_cell In [input_global_verkauf]

            If Application.Intersect(my_cell.Offset(-1, 0), [input_global_verkauf]) Is Nothing _
               And Application.Intersect(my_cell.Offset(0, -1), [input_global_verkauf]) Is Nothing _
               And l_steps < fnBaNumber Then

                l_steps = l_steps + 1

                [set_global_commerce].Copy
                my_cell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                my_cell.PasteSpecial Paste:=xlPasteFormats


                [set_global_miete].Copy
                my_cell.Offset(10, -7).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                my_cell.Offset(10, -7).PasteSpecial Paste:=xlPasteFormats

                [set_global_miete_3_garages].Copy
                my_cell.Offset(-6, -7).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                my_cell.Offset(-6, -10).PasteSpecial Paste:=xlPasteFormats

                [set_total_miete_und_kaufpreis_faktor].Copy
                my_cell.Offset(10, 0).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
                my_cell.Offset(10, 0).PasteSpecial Paste:=xlPasteFormats
                my_cell.Offset(10, 2).FormulaR1C1 = "=12*R[1]C*(R[-18]C[-7]*RC[-5]+R[-17]C[-21]*R[1]C[-5]+R[-16]C[-21]*R[-16]C[-5]+R[-15]C[-21]*R[-15]C[-5]+R[-14]C[-21]*R[-14]C[-5])"

            End If
        Next my_cell
    End If


    Application.CutCopyMode = False

    Call OnEnd
    tbl_Input.Cells(1, 1).Select

End Sub

Public Sub BlackOutRange(Optional b_White_Out As Boolean = False)

    Dim r_range             As Range
    Dim r_range_2           As Range

    Set r_range = Union([input_ba_wohnflache], [input_ba_vorverkaufsquote], [input_ba_w_phasen], _
                        [input_ba_vertriebsstart], [input_ba_verkaufspreis_pro_m2], _
                        [input_ba_anzahl_wohneinheiten], [input_ba_raten_w], [input_ba_c_phasen], _
                        [input_ba_anzahl_gewerbe], [input_ba_raten_c], [input_ba_total_ba], [input_ba_total_gewerbe])

    If b_White_Out Then

        Call DeleteDrawingObjects
        Call DrawBordersAroundRange(True)

    Else

        For Each r_range_2 In r_range.Areas
            Call CoverRange(r_range_2)
        Next r_range_2
        Call DrawBordersAroundRange(False)

    End If

End Sub

Sub CoverRange(ByRef r As Range)

    Dim l As Long, t As Long, w As Long, H As Long

    l = r.Left
    t = r.Top
    w = r.Width
    H = r.Height

    'msoTextOrientationHorizontal
    With ActiveSheet.Shapes
        .AddTextbox(msoTextOrientationVertical, l, t, w, H).Select
        Selection.ShapeRange.line.visible = msoFalse
    End With

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

Public Sub DrawBordersAroundRange(b_remove As Boolean)

    If b_remove Then

        [set_format].Copy
        [input_all_ba].PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        'make the last month white for austria
        If tbl_Input.opt_os Then
            For Each current_cell In [input_construction_time]
                tbl_Input.Cells(current_cell.Row + 8, 12).Font.Color = vbWhite
            Next current_cell
        End If

    Else
        [set_format_without_borders].Copy
        [input_all_ba].PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = xlNone
    End If

End Sub

Public Sub SaveThisWorkbook()
    'We do not have to save the wb each time, but randomly 33% is quite ok.
    'The save is needed, because sometimes the "Steuerelementen" tend to fly around, if we play with them..
    'Currently saved every 2. time

    If make_random(1, 1) = 1 Then
        ThisWorkbook.Save
    End If

End Sub

Public Sub SimpleBorder(r_range As Range)
    
    r_range.Borders.LineStyle = xlNone
    'r_range.Borders(xlDiagonalDown).LineStyle = xlNone
    'r_range.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With r_range.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With r_range.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With r_range.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With r_range.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End Sub
