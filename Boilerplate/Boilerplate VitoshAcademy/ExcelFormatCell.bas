Attribute VB_Name = "ExcelFormatCell"
Option Explicit
Option Private Module

Public Sub FormatAsDate(myCell As Range)
    myCell.NumberFormat = "[$-407]mmm/ yy;@"
End Sub

Public Sub FormatAsPercent(myCell As Range, Optional afterComma = 2)

    If afterComma = 3 Then
        myCell.NumberFormat = "0.000%"
    Else
        myCell.NumberFormat = "0.00%"
    End If

End Sub

Public Sub FormatAsCurrency(myCell As Range, _
                    Optional changeZero = False, _
                    Optional makeGray = True, _
                    Optional makeRound = True)

    Dim isOneCell          As Boolean

    isOneCell = IIf(myCell.Rows.Count + myCell.Columns.Count <> 2, False, True)

    If IsNumeric(myCell.value) And (Not myCell.HasFormula) Then
        myCell.value = Round(myCell.value, 2)
    End If

    If makeRound Then
        myCell.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    Else
        myCell.NumberFormat = "$#,##0.00_);($#,##0.00)"
    End If

    If changeZero Then
        With myCell
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .FormatConditions(1).Font.ThemeColor = xlThemeColorDark1
            .FormatConditions(1).Font.TintAndShade = -0.4
        End With
    End If

    If isOneCell Then
        If makeGray And myCell.value = 0 Then
            With myCell
                .Cells.Font.Color = RGB(191, 191, 191)
            End With
        End If
    End If

End Sub

Public Sub FormatAsEurProM2(myCell As Range)
    myCell.NumberFormat = "#,##0.00 "" € / m²"""
End Sub

Public Sub FormatRedAndBold(myCell As Range, Optional isBold = True)
    myCell.Font.Color = -16777063
    myCell.Font.TintAndShade = 0
    If isBold Then myCell.Font.Bold = True
End Sub

Public Sub WhiteRows(lines As Long, wks As Worksheet)
    
    Dim rowLines As String
    rowLines = lines & ":" & lines
    
    With wks.Rows(rowLines).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
End Sub

Public Sub WhiteCell(myCell As Range)
    myCell.Font.ThemeColor = xlThemeColorDark1
    myCell.Font.TintAndShade = 0
End Sub

Public Sub FormatFontColorToGrey(myCell As Range)
    myCell.Font.Color = RGB(128, 128, 128)
End Sub

