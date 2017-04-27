Public Sub Tdd_CA2()
    
    On Error Resume Next
    
    Dim specs           As New SpecSuite
    Dim myArr           As Variant
    Dim lngSize         As Long: lngSize = 46

    myArr = fnArr_CA0_002
    
    For lngCounter = 0 To UBound(myArr)
    
        lngRow = lngCounter \ lngSize
        lngCol = lngCounter Mod lngSize

        specs.It("CA0_002_F86_Row" & lngRow + 1 & "_Col" & lngCol + 1).Expect(myArr(lngCounter + 1)).ToEqual tbl_calendar.[f86].Offset(lngRow, lngCol).value
        specs.It("MUST_FAIL_CA0_002_F86_Row" & lngRow + 1 & "_Col" & lngCol + 1).Expect(myArr(lngCounter + 1)).ToNotEqual tbl_calendar.[f86].Offset(lngRow, lngCol).value & "1"
        specs.It("MUST_FAIL_CA0_002_F86_Row" & lngRow + 1 & "_Col" & lngCol + 1).Expect(myArr(lngCounter + 1)).ToNotEqual tbl_calendar.[f86].Offset(lngRow, lngCol).value & "2"

    Next lngCounter
    
    InlineRunner.RunSuite specs
    Call specs.TotalTests
    
    On Error GoTo 0
    
End Sub

Public Function fnArr_CA0_002()

    Dim my_arr                  As Variant

    ReDim my_arr(414)
    
    my_arr(1) = 1
    my_arr(2) = 2
    my_arr(413) = 8059.23
    my_arr(414) = 0
    
    fnArr_CA0_002 = my_arr
    
End Function

Public Sub MakeAllValues()

    Dim my_cell                 As Range
    Dim l_counter               As Long
    Dim str                     As String
    Dim str_result              As String
    
    STR_ERROR_REPORT = ""

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
            If my_cell.HasFormula Then
                str = str & """"""
            Else
                str = str & 0
            End If
        End If
        
        If Len(str_result) = 0 Then
            str_result = str
        Else
            str_result = str_result & vbCrLf & str
        End If
    Next my_cell
    
    Debug.Print str_result
    Call CreateLogFile(str_result)

End Sub

Public Sub MakeColorsAllValues()
    
    Dim myCell                  As Range
    Dim lngCounter              As Long
    Dim str                     As String
    Dim strResult               As String
        
    STR_ERROR_REPORT = ""
    
    For Each myCell In Selection
        Call Increment(lngCounter)
        str = vbTab & "my_arr(" & lngCounter & ")= "
        str = str & myCell.Interior.Color
                        
        If Len(strResult) = 0 Then
            strResult = str
        Else
            strResult = strResult & vbCrLf & str
        End If
                
    Next myCell
    
    Debug.Print strResult
    Call CreateLogFile(strResult)
    
End Sub

Public Function codify_time(Optional b_make_str As Boolean = False) As String

    If [set_in_production] Then On Error GoTo codify_Error
    
    Dim dbl_01                  As Variant
    Dim dbl_02                  As Variant
    Dim dbl_now                 As Double
    
    dbl_now = Round(Now(), 8)
    
    dbl_01 = Split(CStr(dbl_now), ",")(0)
    dbl_02 = Split(CStr(dbl_now), ",")(1)
    
    codify_time = Hex(dbl_01) & "_" & Hex(dbl_02)
    
    If b_make_str Then codify_time = "\" & codify_time & ".txt"
    
    On Error GoTo 0
    Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function
