Option Explicit
Option Private Module

Public Function fnLngLocateValueRow(ByVal strTarget, _
                                    ByRef wksTarget As Worksheet, _
                                    Optional lngCol As Long = 1, _
                                    Optional lngMoreValuesFound As Long = 1, _
                                    Optional blnLookForPart = False) As Long

    Dim lngValuesFound      As Long
    Dim rngLocal            As Range
    Dim rngMyCell           As Range
    
    lngValuesFound = lngMoreValuesFound
    Set rngLocal = wksTarget.Range(wksTarget.Cells(1, lngCol), wksTarget.Cells(Rows.Count, lngCol))

    For Each rngMyCell In rngLocal
        If blnLookForPart Then
            If strTarget = Left(rngMyCell, Len(strTarget)) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueRow = rngMyCell.row
                    Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        Else
            If strTarget = Trim(rngMyCell) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueRow = rngMyCell.row
                    Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        End If
    Next rngMyCell

    fnLngLocateValueRow = -1

End Function

Public Function fnLngLocateValueCol(ByVal strTarget, _
                                    ByRef wksTarget As Worksheet, _
                                    Optional lngRow As Long = 1, _
                                    Optional lngMoreValuesFound As Long = 1, _
                                    Optional blnLookForPart = False) As Long
                                    
    Dim lngValuesFound          As Long
    Dim rngLocal                As Range
    Dim rngMyCell               As Range
    
    lngValuesFound = lngMoreValuesFound
    Set rngLocal = wksTarget.Range(wksTarget.Cells(lngRow, 1), wksTarget.Cells(lngRow, Columns.Count))
    
    For Each rngMyCell In rngLocal
        If blnLookForPart Then
            If strTarget = Left(rngMyCell, Len(strTarget)) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueCol = rngMyCell.row
                    Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        Else
            If strTarget = Trim(rngMyCell) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueCol = rngMyCell.Column
                    Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        End If
    Next rngMyCell
    
    fnLngLocateValueCol = -1

End Function

Public Sub FreezeRow(Optional strWsName As String = "Input", Optional strCellAddress As String = "b5")

    Dim ws              As Worksheet

    Set ws = Worksheets(strWsName)

    ActiveWindow.FreezePanes = False
    Application.Goto ws.Range(strCellAddress)
    ActiveWindow.FreezePanes = True

    Set ws = Nothing

End Sub

Public Sub UnfreezeRows(Optional strWsName As String = "Input")
    
    Dim ws              As Worksheet
    
    Set ws = Worksheets(strWsName)
    
    ActiveWindow.FreezePanes = False
    
End Sub

Public Function fnLngReturnFirstPositionWithValue(varSearchInto As Variant, varValue As Variant) As Long

    Dim lngCounter As Long

    For lngCounter = LBound(varSearchInto) To UBound(varSearchInto)
        If varSearchInto(lngCounter) = varValue Then
            fnLngReturnFirstPositionWithValue = lngCounter
            Exit Function
        End If
    Next lngCounter

End Function

Public Function fnDblSumArray(varMyArray As Variant, _
                                Optional lngLastValuesNotToCalculate As Long = 0) As Double

    Dim lngCounter           As Long
    
    For lngCounter = LBound(varMyArray) To UBound(varMyArray) - lngLastValuesNotToCalculate
        fnDblSumArray = fnDblSumArray + varMyArray(lngCounter)
    Next lngCounter
    
End Function

Public Function fnStrChangeCommas(ByVal varMyValue As Variant) As String
    
    Dim strTemp As String
    
    strTemp = CStr(varMyValue)
    fnStrChangeCommas = Replace(strTemp, ",", ".")
    
End Function

Public Function fnVarBubbleSort(ByRef varTempArray As Variant) As Variant

    Dim varTemp                 As Variant
    Dim lngCounter              As Long
    Dim blnNoExchanges          As Boolean

    Do
        blnNoExchanges = True
        
        For lngCounter = LBound(varTempArray) To UBound(varTempArray) - 1
            If CDbl(varTempArray(lngCounter)) > CDbl(varTempArray(lngCounter + 1)) Then
                blnNoExchanges = False
                varTemp = varTempArray(lngCounter)
                varTempArray(lngCounter) = varTempArray(lngCounter + 1)
                varTempArray(lngCounter + 1) = varTemp
            End If
        Next lngCounter
    
    Loop While Not (blnNoExchanges)
    fnVarBubbleSort = varTempArray

   On Error GoTo 0
   Exit Function
   
End Function

Public Function fnDatGetLastDayOfMonth(ByVal datMyDate As Date) As Date
    
    fnDatGetLastDayOfMonth = DateSerial(Year(datMyDate), Month(datMyDate) + 1, 0)
    
End Function

Public Function fnDatGetFirstDayOfMonth(ByVal datMyDate As Date) As Date

    fnDatGetFirstDayOfMonth = DateSerial(Year(datMyDate), Month(datMyDate), 1)

End Function

Public Function fnDatAddMonths(ByVal datMyDate As Date, ByVal lngMonths As Long) As Date

    fnDatAddMonths = fnDatGetLastDayOfMonth(DateAdd("m", lngMonths, datMyDate))

End Function

Public Function fnDatAddMonthsAndGetFirstDate(ByVal datMyDate As Date, ByVal lngMonth As Long) As Date

    fnDatAddMonthsAndGetFirstDate = fnDatGetFirstDayOfMonth(DateAdd("m", lngMonth, datMyDate))

End Function

Public Function fnLngCalculateYearsFromMonths(lngTotalTerm As Long) As Long

    fnLngCalculateYearsFromMonths = lngTotalTerm \ 12
    If lngTotalTerm Mod 12 Then fnLngCalculateYearsFromMonths = fnLngCalculateYearsFromMonths + 1
    
End Function

Public Function fnBlnIsArrayAllocated(varArr As Variant) As Boolean

    On Error Resume Next
    
    fnBlnIsArrayAllocated = IsArray(varArr) And Not IsError(LBound(varArr, 1)) And LBound(varArr, 1) <= UBound(varArr, 1)
    
    On Error GoTo 0

End Function

Public Function fnBlnZeroOrEmpty(ByRef rngCell As Range, Optional blnIsRange = False) As Boolean
    
    Dim rngCurrentCell As Range
    
    If blnIsRange Then
        
        For Each rngCurrentCell In rngCell
            If (IsEmpty(rngCurrentCell) Or rngCurrentCell.value = 0) Then
                fnBlnZeroOrEmpty = True
                Exit Function
            Else
                fnBlnZeroOrEmpty = False
            End If
        Next rngCurrentCell
        
    Else
        If (IsEmpty(rngCell) Or rngCell.value = 0) Then
            fnBlnZeroOrEmpty = True
        Else
            fnBlnZeroOrEmpty = False
        End If
    End If

End Function

Public Function fnLngMillionsEur(ByVal lngMyValue As Long) As Long
    
    fnLngMillionsEur = lngMyValue / 1000000

End Function

Public Function fnDblSumRange(rngRange As Range) As Double

    Dim rngCell As Range

    fnDblSumRange = 0
    
    For Each rngCell In rngRange
        fnDblSumRange = fnDblSumRange + rngRange.value
    Next

End Function

Public Function fnLngMakeRandom(lngDown As Long, lngUp As Long) As Long

    fnLngMakeRandom = CLng((lngUp - lngDown + 1) * Rnd + lngDown)

End Function

Public Function fnBlnCheckIfHidden(rngRange As Range) As Boolean

    If rngRange.EntireRow.Hidden Or rngRange.EntireColumn.Hidden Then
        fnBlnCheckIfHidden = True
    End If

End Function

Function fnLngLastRow(Optional strSheet As String, Optional lngColumnToCheck As Long = 1) As Long
    
    Dim wksSheet             As Worksheet
    
    If strSheet = vbNullString Then
        Set wksSheet = ActiveSheet
    Else
        Set wksSheet = Worksheets(strSheet)
    End If
    
    fnLngLastRow = wksSheet.Cells(wksSheet.Rows.Count, lngColumnToCheck).End(xlUp).row

End Function

Function fnLngLastColumn(Optional strSheet As String, Optional lngRowToCheck As Long = 1) As Long

    Dim wksSheet             As Worksheet

    If strSheet = vbNullString Then
        Set wksSheet = ActiveSheet
    Else
        Set wksSheet = Worksheets(strSheet)
    End If

    fnLngLastColumn = wksSheet.Cells(lngRowToCheck, wksSheet.Columns.Count).End(xlToLeft).Column

End Function

Public Function fnStrLetterCol(ByVal lngCol As Long) As String

    fnStrLetterCol = Split(Cells(1, lngCol).Address, "$")(1)

End Function

Public Function fnBlnValueInArray(ByRef varMyValue As Variant, _
                                  ByRef varMyArray As Variant, _
                                  Optional blnIsString As Boolean = False) As Boolean
                
    Dim lngCounter          As Long

    If blnIsString Then
        varMyArray = Split(varMyArray, ":")
    End If

    For lngCounter = LBound(varMyArray) To UBound(varMyArray)
        varMyArray(lngCounter) = CStr(varMyArray(lngCounter))
    Next lngCounter

    fnBlnValueInArray = Not IsError(Application.Match(CStr(varMyValue), varMyArray, 0))
    
End Function

Public Function fnStrVarRgb2HtmlColor(R As Byte, G As Byte, B As Byte) As String

    'INPUT: Numeric (Base 10) Values for R, G, and B)
    'RETURNS:
    'A string that can be used as an HTML Color
    '(i.e., "#" + the Hexadecimal equivalent)
    'For VBA the RGB is reversed. R and B are revered...

    Dim varHexR         As Variant
    Dim varHexB         As Variant
    Dim varHexG         As Variant

    On Error GoTo ErrorHandler

    'R
    varHexR = Hex(R)
    If Len(varHexR) < 2 Then varHexR = "0" & varHexR

    'Get Green Hex
    varHexG = Hex(G)
    If Len(varHexG) < 2 Then varHexG = "0" & varHexG

    varHexB = Hex(B)
    If Len(varHexB) < 2 Then varHexB = "0" & varHexB


    fnStrVarRgb2HtmlColor = "#" & varHexR & varHexG & varHexB
ErrorHandler:

End Function

Function fnBlnNamedRngExists(strRangeName As String) As Boolean

    Dim strMyRange As Range

    On Error Resume Next

    Set strMyRange = Range(strRangeName)
    If Not strMyRange Is Nothing Then fnBlnNamedRngExists = True

    On Error GoTo 0

End Function

Function fnStrGetRgb2(lngLong) As String

    Dim R As Long
    Dim G As Long
    Dim B As Long

    R = lngLong Mod 256
    G = lngLong \ 256 Mod 256
    B = lngLong \ 65536 Mod 256
    fnStrGetRgb2 = "R=" & R & ", G=" & G & ", B=" & B
    
End Function

Sub ChangeAllNames()
    
   If [set_in_production] Then On Error GoTo ChangeAllNames_Error

    Dim lngCounter          As Long
    Dim strOld              As String
    Dim strNew              As String
    
    For lngCounter = 1 To ActiveWorkbook.Names.Count

        If InStr(1, ActiveWorkbook.Names(lngCounter), "old", vbTextCompare) Then
            strOld = ActiveWorkbook.Names(lngCounter).RefersToR1C1
            strOld = Replace(strOld, "old", "")
            Debug.Print strNew
            
            With ActiveWorkbook.Names(ActiveWorkbook.Names(lngCounter).Name)
                .RefersToR1C1 = strNew

            End With
        End If
    Next lngCounter

    On Error GoTo 0
    Exit Sub

ChangeAllNames_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure change_all_names of Sub mod_StandardSubs"

End Sub

Public Sub DeleteCommentInSelection()
    
   If [set_in_production] Then On Error GoTo DeleteCommentInSelection_Error

    Dim rngCurrentCell As Range
    
    For Each rngCurrentCell In Selection
        rngCurrentCell.ClearComments
    Next rngCurrentCell
    
   On Error GoTo 0
   Exit Sub

DeleteCommentInSelection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DeleteCommentInSelection of Sub mod_StandardSubs"

End Sub

Public Sub SelectMeA1RangeEverywhere()

    If [set_in_production] Then On Error GoTo SelectMeA1RangeEverywhere_Error

    Dim wksSheet As Worksheet

    For Each wksSheet In ThisWorkbook.Sheets
        If wksSheet.Visible = xlSheetVisible Then
            wksSheet.Activate
            wksSheet.Cells(1, 1).Select
        End If
    Next wksSheet
    
    Worksheets(1).Select

    On Error GoTo 0
    Exit Sub

SelectMeA1RangeEverywhere_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SelectMeA1RangeEverywhere of Sub mod_StandardSubs"

End Sub

Sub HideShowComments(Optional bShowComments As Boolean = False, Optional rngMyRange As Range = Nothing)
    
    Dim rngCurrentCell    As Range
    
    If [set_in_production] Then On Error GoTo HideShowComments_Error
    If rngMyRange Is Nothing Then Set rngMyRange = Range("A1:AO1000")
        
    For Each rngCurrentCell In rngMyRange
        If Not rngCurrentCell.Comment Is Nothing Then
            rngCurrentCell.Comment.Visible = bShowComments
        End If
    Next rngCurrentCell

    On Error GoTo 0
    Exit Sub

HideShowComments_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HideShowComments of Sub mod_StandardSubs"

End Sub

Public Sub Info()
    
    If Not fnBlnValueInArray(Environ("Username"), ADMINS, True) Then
        Debug.Print "no"
        Exit Sub
    End If

    Call UnhideAll 'UnprotectAll is included
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    Debug.Print "Done."

    Call EnableMySaves

End Sub

Public Sub EnableMySaves()

    Application.OnKey "%{F11}"
    Application.OnKey "^c"
    Application.OnKey "^v"
    Application.OnKey "^x"

End Sub

Public Sub DisabledCombination()
    
    'This is the disabled combination for Application.OnKey

End Sub

Public Sub DisableMySaves()

    Application.OnKey "%{F11}", "DisabledCombination"
    Application.OnKey "^c", "DisabledCombination"
    Application.OnKey "^v", "DisabledCombination"
    Application.OnKey "^x", "DisabledCombination"

End Sub

Public Sub PrintArray(ByRef varMyArray As Variant)

    Dim lngCounter As Long
    
    For lngCounter = LBound(varMyArray) To UBound(varMyArray)
        Debug.Print lngCounter & " --> " & varMyArray(lngCounter)
    Next lngCounter
    
End Sub

Public Sub FormatAsDate(ByRef rngCell As Range)

    rngCell.NumberFormat = "[$-407]mmm/ yy;@"
    
End Sub

Public Sub FormatAsPercent(ByRef rngMyCell As Range, Optional lngNumber = 2)

    If lngNumber = 3 Then
        rngMyCell.NumberFormat = "0.000%"
    Else
        rngMyCell.NumberFormat = "0.00%"
    End If

End Sub

Public Sub FormatAsCurrency(ByRef rngMyCell As Range, _
                            Optional ByVal blnChangeZero = False, _
                            Optional blnMakeGray = True, _
                            Optional blnMakeRound = True)

    Dim blnIsAlone          As Boolean

    blnIsAlone = IIf(rngMyCell.Rows.Count + rngMyCell.Columns.Count <> 2, False, True)

    If IsNumeric(rngMyCell.value) And (Not rngMyCell.HasFormula) Then
        rngMyCell.value = Round(rngMyCell.value, 2)
    End If

    If blnMakeRound Then
        rngMyCell.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Else
        rngMyCell.NumberFormat = "$#,##0.00_);($#,##0.00)"
    End If

    If blnChangeZero Then
        With rngMyCell
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .FormatConditions(1).Font.ThemeColor = xlThemeColorDark1
            .FormatConditions(1).Font.TintAndShade = -0.4
        End With
    End If

    If blnIsAlone Then
        If blnMakeGray And rngMyCell.value = 0 Then
            With rngMyCell
                .Cells.Font.Color = RGB(191, 191, 191)
            End With
        End If
    End If

End Sub

Public Sub FormatAsEurProM2(rngMyCell As Range)

    rngMyCell.NumberFormat = "#,##0.00 "" € / m²"""

End Sub

Public Sub FormatRedAndBold(ByRef rngMyCell As Range, Optional blnIsBold = True)
    
    rngMyCell.Font.Color = -16777063
    rngMyCell.Font.TintAndShade = 0

    If blnIsBold Then rngMyCell.Font.Bold = True
    
End Sub

Public Sub WhiteYourself(ByVal lngLines As Long, ByRef wksMySheet As Worksheet)
    
    Dim strLines                       As String
    strLines = lngLines & ":" & lngLines
    
    With wksMySheet.Rows(strLines).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
End Sub

Public Sub WhiteCell(ByRef rngMyCell As Range)
    
    rngMyCell.Font.ThemeColor = xlThemeColorDark1
    rngMyCell.Font.TintAndShade = 0
    
End Sub

Public Sub FormatFontColorToGrey(ByRef rngMyCell As Range)

    rngMyCell.Font.Color = RGB(128, 128, 128)

End Sub
Public Sub CopyValues(RngSource As Range, RngTarget As Range)
 
    RngTarget.Resize(RngSource.Rows.Count, RngSource.Columns.Count).value = RngSource.value
 
End Sub

Public Sub UnhideAll()

    Dim wksMySheet              As Worksheet

    For Each wksMySheet In ThisWorkbook.Worksheets
       wksMySheet.Visible = xlSheetVisible
    Next wksMySheet

    Call UnprotectAll

End Sub

Public Sub UnprotectAll()

    Dim lngCounter As Long
    
    For lngCounter = ActiveWorkbook.Worksheets.Count To 1 Step -1
        ActiveWorkbook.Worksheets(lngCounter).Unprotect Password:=s_CONST
    Next lngCounter
    
End Sub

Public Sub HideNeeded()
    
    Dim varSheet                    As Variant
    Dim varArrVisibleSheets         As Variant
    Dim varArrHiddenSheets          As Variant

    Call OnStart

    varArrVisibleSheets = Array(tblInput)
    varArrHiddenSheets = Array(tblSettings)

    For Each varSheet In varArrVisibleSheets
        varSheet.Visible = xlSheetVisible
    Next varSheet

    For Each varSheet In varArrHiddenSheets
        varSheet.Visible = xlSheetVeryHidden
    Next varSheet

    Call OnEnd

End Sub

Public Sub AddCommentToSelection(rngMyComment As Range)
    
    Dim blnMyBoolean            As Boolean
    Dim rngCurrentCell            As Range
    
    blnMyBoolean = True
    
    For Each rngCurrentCell In Selection
        If blnMyBoolean Then
            rngCurrentCell.ClearComments
            rngCurrentCell.AddComment rngMyComment.Text
            rngCurrentCell.Comment.Visible = False
            rngCurrentCell.Comment.Shape.ScaleWidth 4, msoFalse, msoScaleFromTopLeft
            rngCurrentCell.Comment.Shape.ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft
        End If
        'blnMyBoolean = Not blnMyBoolean
    Next rngCurrentCell

End Sub

Sub DeleteDrawingObjects()

    Dim lngCounter           As Long

    For lngCounter = tblInput.DrawingObjects().Count To 1 Step -1
        If Left(tblInput.DrawingObjects(lngCounter).Name, 7) = "TextBox" Then
            tblInput.DrawingObjects(lngCounter).Delete
        End If
    Next lngCounter

End Sub

Sub CoverRange(ByRef rngRange As Range)
    
    Dim lngLeft             As Long
    Dim lngTop              As Long
    Dim lngWidth            As Long
    Dim lngHeight           As Long
    
    lngLeft = rngRange.Left
    lngTop = rngRange.Top
    lngWidth = rngRange.Width
    lngHeight = rngRange.Height
    
    With ActiveSheet.Shapes
        .AddTextbox(msoTextOrientationVertical, lngLeft, lngTop, lngWidth, lngHeight).Select
        Selection.ShapeRange.Line.Visible = msoFalse
    End With

End Sub

Public Sub PrintPDF(Optional blnBlack As Boolean = False, _
                    Optional rngInputPrintArea As Range = Nothing, _
                    Optional rngInputObjectAddress As Range = Nothing, _
                    Optional rngInputCalculationDate As Range = Nothing)

    On Error GoTo PrintPDF_Error

    If rngInputPrintArea Is Nothing Then Set rngInputPrintArea = [input_print_area]
    If rngInputObjectAddress Is Nothing Then Set rngInputObjectAddress = [input_object_address]
    If rngInputCalculationDate Is Nothing Then Set rngInputCalculationDate = [input_calculation_date]

    ActiveSheet.PageSetup.Zoom = False
    ActiveSheet.PageSetup.BlackAndWhite = blnBlack

    rngInputPrintArea.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=CStr(rngInputObjectAddress & "_" & rngInputCalculationDate), _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True

    On Error GoTo 0
    Exit Sub

PrintPDF_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPDF of Modul mod_Drucken"

End Sub

Public Sub PrintPage(Optional blnBlack As Boolean = False)

    Dim wksSheet                    As Worksheet
    Dim rngPrint                    As Range
    Dim strReducePaperTitle         As String

    On Error GoTo PrintPage_Error

    strReducePaperTitle = "Reduzieren Sie den Papierverbrauch"
    ActiveSheet.PageSetup.BlackAndWhite = blnBlack

    Set wksSheet = ActiveSheet
    Set rngPrint = [input_print_area]

    With wksSheet.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With

    Select Case MsgBox("Sind Sie sicher, dass Sie drucken moechten?", vbYesNo Or vbQuestion Or vbDefaultButton1, strReducePaperTitle)

    Case vbYes
        Select Case MsgBox("Wirklich sicher, dass Sie drucken moechten?", vbYesNo Or vbQuestion Or vbDefaultButton1, strReducePaperTitle)
        Case vbYes
            rngPrint.PrintOut
        End Select
    End Select

    On Error GoTo 0
    Exit Sub

PrintPage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPage of Modul mod_Drucken"

End Sub

Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    
    Application.StatusBar = False
    
End Sub

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    
    ActiveWindow.View = xlNormalView

End Sub

Public Sub DeleteName(strName As String)

    On Error GoTo DeleteName_Error

    ActiveWorkbook.Names(strName).Delete
    Debug.Print strName & " is deleted!"
    
    On Error GoTo 0
    Exit Sub

DeleteName_Error:

    Debug.Print strName & " not present or some error"
    On Error GoTo 0
    
End Sub

Public Sub LockScroll(ByRef varMyArray As Variant)
    
    Dim lngCounter           As Long
    
    If Not Len(Join(varMyArray)) > 0 Then Exit Sub
    
    For lngCounter = 0 To UBound(varMyArray) Step 2
        ThisWorkbook.Sheets(varMyArray(lngCounter)).ScrollArea = varMyArray(lngCounter + 1)
    Next lngCounter
    
End Sub

Public Sub Decrement(ByRef VarValueToDecrement As Variant, Optional dblMinus As Double = 1)

    VarValueToDecrement.value = VarValueToDecrement - dblMinus

End Sub

Public Sub Increment(ByRef VarValueToIncrement As Variant, Optional dblPlus As Double = 1)

    VarValueToIncrement = VarValueToIncrement + dblPlus

End Sub

Sub StyleKiller()

    Dim styStyle                As Style
    
    For Each styStyle In ThisWorkbook.Styles
        If Not styStyle.BuiltIn Then
            Debug.Print styStyle.Name
            styStyle.Delete
        End If
    Next styStyle

End Sub


