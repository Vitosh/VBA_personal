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
        myArray(cnt) = CStr(myArray(cnt))
    Next cnt

    valueInArray = Not IsError(Application.Match(CStr(myValue), myArray, 0))

End Function
