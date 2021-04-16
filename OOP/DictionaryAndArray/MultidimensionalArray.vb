Sub PrintMultidimensionalArrayExample()

    Dim myRange As Range
    Set myRange = Range("BB1:BE9")
    
    Dim myArray As Variant
    myArray = myRange
    
    Debug.Print UBound(myArray, 1)  'count of excel cells in a column
    Debug.Print UBound(myArray, 2)  'count of excel cells in a row
    
    Debug.Print LBound(myArray, 1)  'index of first cell in column
    Debug.Print LBound(myArray, 2)  'index of first cell in row
    
    PrintArray GetRowFromMdArray(myArray, 1)
    PrintArray GetColumnFromMdArray(myArray, UBound(myArray, 2))

End Sub

Function GetColumnFromMdArray(myArray As Variant, myCol As Long) As Variant
    
    'returning a column from multidimensional array
    'the returned array is 0-based, but the 0th element is Empty.
    
    Dim i As Long
    Dim result As Variant
    Dim size As Long: size = UBound(myArray, 1)
    ReDim result(size)
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        result(i) = myArray(i, myCol)
    Next
    
    GetColumnFromMdArray = result
    
End Function

Function GetRowFromMdArray(myArray As Variant, myRow As Long) As Variant
    
    'returning a row from multidimensional array
    'the returned array is 0-based, but the 0th element is Empty.
    
    Dim i As Long
    Dim result As Variant
    Dim size As Long: size = UBound(myArray, 2)
    ReDim result(size)
    
    For i = LBound(myArray, 2) To UBound(myArray, 2)
        result(i) = myArray(myRow, i)
    Next
    
    GetRowFromMdArray = result
    
End Function

Public Sub PrintArray(myArray As Variant)

    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        Debug.Print i & " --> " & myArray(i)
    Next i
    
End Sub

Public Function GetIndexInArrayFirstLast(myArray As Variant, myValue As String, Optional firstNeeded As Boolean = True) As Long
    
    GetIndexInArrayFirstLast = GENERAL_NUMBERS.MINUS_ONE
    
    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        If Trim(UCase(myArray(i))) = Trim(UCase(myValue)) Then
            GetIndexInArrayFirstLast = i
            If firstNeeded Then Exit Function
        End If
    Next
    
End Function
