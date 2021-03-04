Public Sub CloseAllExcelFilesExceptCurrent()

    Dim wb As Workbook
    
    Application.ScreenUpdating = False
    
    For Each wb In Workbooks

        If Not wb.ReadOnly Then wb.Save
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close
        End If
    Next wb
    
End Sub


Public Function valueInArray(myValue As Variant, myArray As Variant) As Boolean

    Dim cnt As Long

    For cnt = LBound(myArray) To UBound(myArray)
        If LCase(CStr(myValue)) = LCase(CStr(myArray(cnt))) Then
            valueInArray = True
            Exit Function
        End If
    Next cnt

End Function

Sub CheckUser()

    Dim userNames As Variant
    userNames = Array("User1", "User2", "User3")

    If valueInArray(Environ("UserName"), userNames) Then
        Debug.Print "User Present"
    Else
        Debug.Print "User Not Present"
    End If
    
End Sub


Sub ChangeTheFont(lookFor As String, currentRange As Range, myColor As Long)

    Dim startPosition As Long: startPosition = InStr(1, currentRange.Value2, lookFor)
    Dim endPosition As Long: endPosition = startPosition + Len(currentRange.Value2)

    With currentRange.Characters(startPosition, Len(lookFor)).Font
        .Color = myColor
        .Bold = True
    End With
End Sub

Public Function PositionInArray(myValue As Variant, myArray As Variant, Optional timesSeenBefore = 0) As Long
    
    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        If Trim(myValue) = Trim(myArray(i)) Then
            If timesSeenBefore = 0 Then
                PositionInArray = i
                Exit Function
            Else
                timesSeenBefore = timesSeenBefore - 1
            End If
        End If
    Next
    
    PositionInArray = -1
    
End Function
