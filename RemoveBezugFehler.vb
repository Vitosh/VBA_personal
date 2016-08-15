Option Explicit
    
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

