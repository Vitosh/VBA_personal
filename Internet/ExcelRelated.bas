Attribute VB_Name = "ExcelRelated"
Option Explicit

Public Function GetNextKeyWord() As String
    
    With tblInput
        Dim lastRowB As Long
        lastRowB = lastRow(.Name, 2) + 1
        GetNextKeyWord = Trim(.Cells(lastRowB, 1))
        If Len(GetNextKeyWord) <> 0 Then .Cells(lastRowB, 2) = Now
    End With
    
End Function

Public Sub WriteFormulas()
    
    Dim i As Long
    With tblInput
        For i = lastRow(.Name) To 2 Step -1
            .Cells(i, 3).FormulaR1C1 = "=COUNTIF(Summary!C[1],Input!RC[-2])"
            
            .Cells(i, 4).FormulaArray = "=MAX(IF(Summary!C=RC[-3],Summary!C[-1]))"
            FormatUSD .Cells(i, 4)
            
            .Cells(i, 5).FormulaArray = "=AVERAGE(IF(Summary!C[-1]=Input!RC[-4],Summary!C[-2]))"
            FormatUSD .Cells(i, 5)
        Next i
    End With
    
End Sub

Public Sub FixWorksheets()
    OnStart
    With tblInput
        .Range("B1") = "Start Time"
        .Range("C1") = "Count"
        .Range("D1") = "Max"
        .Range("E1") = "Average"
    End With
    
    With tblSummary
        .Range("A1") = "Title"
        .Range("B1") = "Author"
        .Range("C1") = "Price"
        .Range("D1") = "Keyword"
    End With
    
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Columns.AutoFit
    Next ws
    OnEnd
End Sub

Public Sub FormatUSD(myRange As Range)

    myRange.NumberFormat = "_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* ""-""??_ ;_-@_ "

End Sub

        
Public Sub CleanWorksheets()

    tblRawData.Cells.Delete
    tblSummary.Cells.Delete
    tblInput.Columns("B:F").Delete
        
End Sub

Public Function GetNthString(n As Long, myRange As Range) As String
    
    Dim i As Long
    Dim myVar As Variant
    
    myVar = Split(myRange, vbCrLf)
    For i = LBound(myVar) To UBound(myVar)
        If Len(myVar(i)) > 0 And n = 0 Then
            GetNthString = myVar(i)
            Exit Function
        ElseIf Len(myVar(i)) > 0 Then
            n = n - 1
        End If
    Next i
    
End Function


Public Function GetPrice(myRange As Range) As String
    
    Dim i As Long
    Dim myVar As Variant
    myVar = Split(myRange, "$")
    
    If UBound(myVar) > 0 Then
        GetPrice = Mid(myVar(1), 1, InStr(1, myVar(1), " "))
    Else
        GetPrice = ""
    End If
        
End Function

Public Sub WriteToExcel(appIE As Object, keyword As String)

    If IN_PRODUCTION Then On Error GoTo WriteToExcel_Error
    
    Dim allData As Object
    Set allData = appIE.document.getElementById("s-results-list-atf")
    
    Dim book As Object
    Dim myRow As Long
        
    For Each book In allData.getElementsByClassName("a-fixed-left-grid-inner")
        With tblRawData
            myRow = lastRow(.Name) + 1
            On Error Resume Next
            .Cells(myRow, 1) = book.innertext
            .Cells(myRow, 2) = keyword
            On Error GoTo 0
        End With
    Next
        
    IeErrors = 0
    
    On Error GoTo 0
    Exit Sub

WriteToExcel_Error:

    IeErrors = IeErrors + 1
    
    If IeErrors > MAX_IE_ERRORS Then
        Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteToExcel, line " & Erl & "."
    Else
        LogMe "WriteToExcel", IeErrors, keyword, IeErrors
        WriteToExcel appIE, keyword
    End If
    
End Sub

Public Sub RawDataToStructured(keyword As String, firstRow As Long)
    
    Dim i As Long
    For i = firstRow To lastRow(tblRawData.Name)
        With tblRawData
            If InStr(1, .Cells(i, 1), "Sponsored ") < 1 Then
                Dim title As String
                title = GetNthString(0, .Cells(i, 1))
                Dim author As String
                author = GetNthString(1, .Cells(i, 1))
                Dim price As String
                price = GetPrice(.Cells(i, 1))
                If Not IsNumeric(price) Or price = "0" Then price = ""
                Dim currentRow As String: currentRow = lastRow(tblSummary.Name) + 1
                With tblSummary
                    .Cells(currentRow, 1) = title
                    .Cells(currentRow, 2) = author
                    .Cells(currentRow, 3) = price
                    .Cells(currentRow, 4) = keyword
                End With
            End If
        End With
    Next i

End Sub

Public Function lastRow(wsName As String, Optional columnToCheck As Long = 1) As Long

    Dim ws As Worksheet
    Set ws = Worksheets(wsName)
    lastRow = ws.Cells(ws.Rows.Count, columnToCheck).End(xlUp).Row

End Function


