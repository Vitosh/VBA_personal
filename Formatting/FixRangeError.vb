Sub ErrorInFormulas()

    'Formatting condition, conditional formatting, external

    Dim ws As Worksheet, r As Range
    Dim cf As FormatCondition

    For Each ws In Worksheets
    
        For Each r In ws.UsedRange
            If IsError(r) Then
                Debug.Print r.Parent.Name, r.Address, r.Formula
            End If
        Next
        
        For Each cf In ws.Cells.FormatConditions
            Debug.Print cf.AppliesTo.Address, cf.Type, cf.Formula1, cf.Interior.COLOR, cf.Font.Name, ws.Name
        Next
    Next
    
End Sub

Sub ListAllConditionalFormatting()

    Dim cf As FormatCondition
    Dim ws As Worksheet
    Set ws = ActiveSheet
    For Each cf In ws.Cells.FormatConditions
        Debug.Print cf.AppliesTo.Address, cf.Type, cf.Formula1, cf.Interior.COLOR, cf.Font.Name
    Next cf

End Sub


Sub ErrorList()

    Dim ws As Worksheet
    Dim rng1 As Range
    Dim strOut As String
    
    For Each ws In ThisWorkbook.Worksheets
        Set rng1 = Nothing
        On Error Resume Next
        Set rng1 = ws.Cells.SpecialCells(xlFormulas, xlErrors)
        On Error GoTo 0
        If Not rng1 Is Nothing Then strOut = strOut & (ws.Name & " has " & rng1.Cells.count & " errors" & vbNewLine)
    Next ws
    
    If Len(strOut) > 0 Then
        Debug.Print "Error List:" & vbNewLine & strOut
    Else
        Debug.Print "No Errors"
    End If
    
End Sub

    
Sub FixRangeError()
    
    On Error GoTo FixRangeError_Error

        Dim r_range         As Range
        Dim str_text        As String
        Dim l_counter       As Long
        Dim str_result      As String
        
        Dim arr_result      As Variant
        Dim arr_range       As Variant
        
        For Each r_range In ActiveSheet.UsedRange
			str_text = ""
            If r_range.HasFormula Then
                ReDim arr_result(0)
                str_text = Replace(r_range.Formula, "=", "")
                
                arr_range = Split(str_text, "+")
                
                For l_counter = LBound(arr_range) To UBound(arr_range)
                    If Not InStr(arr_range(l_counter), "#") > 0 Then
                        ReDim Preserve arr_result(UBound(arr_result) + 1)
                        arr_result(UBound(arr_result)) = arr_range(l_counter)
                    End If
                Next l_counter
                
                For l_counter = LBound(arr_result) + 1 To UBound(arr_result)
                    str_result = str_result & "+" & arr_result(l_counter)
                Next l_counter
                
                str_result = "=" & Right(str_result, Len(str_result) - 1)
                
                r_range.Formula = str_result
            End If
        Next r_range
                

   On Error GoTo 0
   Exit Sub

FixRangeError_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FixRangeError of Sub Modul1"

End Sub

'---------------------------------------------------------------------------------------
' Method : FindMeTheCellWithError
' Author : v.doynov
' Date   : 01.09.2017
' Purpose: Show the errors. Print the errors in a worksheet. Look for errors. Search errors.
'---------------------------------------------------------------------------------------
Public Sub FindMeTheCellWithError()

    Dim rngCell     As Range
    Dim wks         As Worksheet

    For Each wks In ThisWorkbook.Worksheets
        For Each rngCell In wks.UsedRange
            If IsError(rngCell) Then
                Debug.Print rngCell.Address
                Debug.Print rngCell.Parent.name
            End If
        Next rngCell
    Next wks

End Sub
