'https://www.vitoshacademy.com/vba-vlookup-with-multiple-criteria-in-excel-without-excel-formula-but-with-vba/

Function GetLookupDataTriple(wks As Worksheet, tableName As String, lookIntoColumn As String, myArray As Variant) As Variant
    
    Dim lo As ListObject
    Set lo = wks.ListObjects(tableName)
    
    Dim i As Long
    For i = 2 To lo.ListColumns(myArray(0)).Range.Rows.Count
        If lo.ListColumns(myArray(0)).Range.Cells(RowIndex:=i) = myArray(1) Then
            If lo.ListColumns(myArray(2)).Range.Cells(RowIndex:=i) = myArray(3) Then
                If lo.ListColumns(myArray(4)).Range.Cells(RowIndex:=i) = myArray(5) Then
                    GetLookupDataTriple = lo.ListColumns(lookIntoColumn).Range.Cells(RowIndex:=i)
                    Exit Function
                End If
            End If
        End If
    Next i
    
    GetLookupDataTriple = -1
    
End Function

Function GetLookupDataDouble(wks As Worksheet, tableName As String, lookIntoColumn As String, myArray As Variant) As Variant
    
    Dim lo As ListObject
    Set lo = wks.ListObjects(tableName)
    
    Dim i As Long
    For i = 2 To lo.ListColumns(myArray(0)).Range.Rows.Count
        If lo.ListColumns(myArray(0)).Range.Cells(RowIndex:=i) = myArray(1) Then
            If lo.ListColumns(myArray(2)).Range.Cells(RowIndex:=i) = myArray(3) Then
                GetLookupDataDouble = lo.ListColumns(lookIntoColumn).Range.Cells(RowIndex:=i)
                Exit Function
            End If
        End If
    Next i
    
    GetLookupDataDouble = -1
    
End Function
