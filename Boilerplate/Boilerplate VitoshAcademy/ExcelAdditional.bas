Attribute VB_Name = "ExcelAdditional"
Option Explicit
Option Private Module

Public Sub FreezeRow(Optional wsName As String = "Input", Optional cellAddress As String = "b5")

    Dim ws As Worksheet
    Set ws = Worksheets(wsName)

    ActiveWindow.FreezePanes = False
    Application.Goto ws.Range(cellAddress)
    ActiveWindow.FreezePanes = True

End Sub

Public Sub UnfreezeRows(Optional wsName As String = "Input")
    
    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    ActiveWindow.FreezePanes = False
    
End Sub

Public Function SumArray(myArray As Variant, Optional lastValuesNotToCalculate As Long = 0) As Double

    Dim i As Long
    
    For i = LBound(myArray) To UBound(myArray) - lastValuesNotToCalculate
        SumArray = SumArray + myArray(i)
    Next
    
End Function

Public Function ChangeCommas(ByVal myValue As Variant) As String
    
    Dim temp As String
    
    temp = CStr(myValue)
    ChangeCommas = Replace(temp, ",", ".")
    
End Function

Public Function BubbleSort(ByRef myArray As Variant) As Variant

    Dim temp As Variant
    Dim i As Long
    Dim noExchanges As Boolean

    Do
        noExchanges = True
        
        For i = LBound(myArray) To UBound(myArray) - 1
            If CDbl(myArray(i)) > CDbl(myArray(i + 1)) Then
                noExchanges = False
                temp = myArray(i)
                myArray(i) = myArray(i + 1)
                myArray(i + 1) = temp
            End If
        Next i
    
    Loop While Not (noExchanges)
    
    BubbleSort = myArray

    On Error GoTo 0
    Exit Function
   
End Function

Public Function IsArrayAllocated(varArr As Variant) As Boolean

    On Error Resume Next
    IsArrayAllocated = IsArray(varArr) And Not IsError(LBound(varArr, 1)) And LBound(varArr, 1) <= UBound(varArr, 1)
    On Error GoTo 0

End Function

Public Function RangeIsZeroOrEmpty(myRange As Range) As Boolean
    
    Dim myCell As Range
    
    If myRange.Cells.Count > 1 Then
        
        For Each myCell In myRange
            If (IsEmpty(myCell) Or myCell.value = 0) Then
                RangeIsZeroOrEmpty = True
                Exit Function
            Else
                RangeIsZeroOrEmpty = False
            End If
        Next myCell
    Else
        If (IsEmpty(myRange) Or myRange.value = 0) Then
            RangeIsZeroOrEmpty = True
            Exit Function
        Else
            RangeIsZeroOrEmpty = False
        End If
    End If

End Function

Public Function MakeRandom(lowest As Long, highest As Long) As Long

    MakeRandom = CLng((highest - lowest) * Rnd + lowest)

End Function

Public Function IsRangeHidden(myRange As Range) As Boolean

    If myRange.EntireRow.Hidden Or myRange.EntireColumn.Hidden Then
        IsRangeHidden = True
    End If

End Function

Public Function ColumnNumberToLetter(col As Long) As String
    ColumnNumberToLetter = Split(Cells(1, col).Address, "$")(1)
End Function

Public Function IsValueInArray(varMyValue As Variant, myArray As Variant, _
                                            Optional isValueString As Boolean = False) As Boolean
                
    Dim i As Long

    If isValueString Then
        myArray = Split(myArray, ":")
    End If

    For i = LBound(myArray) To UBound(myArray)
        myArray(i) = CStr(myArray(i))
    Next i

    IsValueInArray = Not IsError(Application.Match(CStr(varMyValue), myArray, 0))
    
End Function

Public Function Rgb2HtmlColor(r As Byte, g As Byte, b As Byte) As String

    'INPUT: Numeric (Base 10) Values for R, G, and B)
    'RETURNS:
    'A string that can be used as an HTML Color
    '(i.e., "#" + the Hexadecimal equivalent)
    'For VBA the RGB is reversed. R and B are revered...

    Dim varHexR         As Variant
    Dim varHexB         As Variant
    Dim varHexG         As Variant

    'R
    varHexR = Hex(r)
    If Len(varHexR) < 2 Then varHexR = "0" & varHexR

    'Get Green Hex
    varHexG = Hex(g)
    If Len(varHexG) < 2 Then varHexG = "0" & varHexG

    varHexB = Hex(b)
    If Len(varHexB) < 2 Then varHexB = "0" & varHexB


    Rgb2HtmlColor = "#" & varHexR & varHexG & varHexB
    
End Function

Function NamedRangeExists(rangeName As String) As Boolean

    On Error Resume Next
    
    Dim myRange As Range
    Set myRange = Range(rangeName)
    If Not myRange Is Nothing Then NamedRangeExists = True

    On Error GoTo 0

End Function

Function GetRgb(lngLong) As String

    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = lngLong Mod 256
    g = lngLong \ 256 Mod 256
    b = lngLong \ 65536 Mod 256
    GetRgb = "R=" & r & ", G=" & g & ", B=" & b
    
End Function

Public Sub CopyValues(mySource As Range, myTarget As Range)
    myTarget.Resize(mySource.Rows.Count, mySource.Columns.Count).value = mySource.value
End Sub

Public Sub OnEnd()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True

    ActiveWindow.View = xlNormalView
    Application.StatusBar = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False
    
End Sub

Public Sub OnStart()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    
    ActiveWindow.View = xlNormalView
    Application.StatusBar = False
    Application.Calculation = xlAutomatic
    ThisWorkbook.Date1904 = False

End Sub
