'sort array arraysort array sort sortlist listsort sortlist bubblesort bubble sort

Option Explicit

Public Const STR_SPACE = "-" & vbTab

Public Function fnVarBubbleSort(ByRef varTempArray As Variant) As Variant

    Dim varTemp                 As Variant
    Dim lngCounter              As Long
    Dim blnNoExchanges          As Boolean

    Do
        blnNoExchanges = True
        
        For lngCounter = LBound(varTempArray) To UBound(varTempArray) - 1
            If CDbl(varTempArray(lngCounter)) > CDbl(varTempArray(lngCounter + 1)) Then
                blnNoExchanges = False
                varTemp = varTempArray(lngCounter)
                varTempArray(lngCounter) = varTempArray(lngCounter + 1)
                varTempArray(lngCounter + 1) = varTemp
            End If
        Next lngCounter
    
    Loop While Not (blnNoExchanges)
    fnVarBubbleSort = varTempArray

   On Error GoTo 0
   Exit Function
   
End Function

Public Function fnListToArray(ByRef myList As Collection) As Variant
    
    Dim lngCounter  As Long
    Dim myVar       As Variant
    
    ReDim myVar(myList.Count)
    
    For lngCounter = 0 To myList.Count - 1
        myVar(lngCounter) = myList(lngCounter + 1)
    Next lngCounter
    
    fnListToArray = myVar
    
End Function

Public Function fnArrayToList(ByRef myArray As Variant) As Collection

    Dim lngCounter  As Long
    Dim myCol       As New Collection
    
    For lngCounter = LBound(myArray) To UBound(myArray)
        myCol.Add myArray(lngCounter)
    Next lngCounter
    
    Set fnArrayToList = myCol

End Function


Public Sub TestMe()

    Dim colCollection   As New Collection
    Dim varElement      As Variant

    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.01.2012")
    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.05.2015")
    colCollection.Add CDate("01.01.2016")
    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.01.2011")
    colCollection.Add CDate("01.01.2011")

    Set colCollection = fnArrayToList(fnVarBubbleSort(fnListToArray(colCollection)))

    For Each varElement In colCollection
        Debug.Print varElement
    Next varElement

End Sub
