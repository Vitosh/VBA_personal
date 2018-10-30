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
        If CStr(myValue) = CStr(myArray(cnt)) Then
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

