Option Explicit

'---------------------------------------------------------------------------------------
' Method : TestMe
' Date   : 22.01.2018
' Purpose: Do not try to sum array like this :)
'           Sample for recursion sum.
'---------------------------------------------------------------------------------------
Public Sub TestMe()
    Debug.Print SumArrayRecursion(Array(1, 2, 4, 8))
End Sub

Public Function SumArrayRecursion(arr As Variant) As Long

    Dim cnt     As Long
    Dim newArr  As Variant
    
    If LBound(arr) = UBound(arr) Then
        SumArrayRecursion = arr(0)
        Exit Function
    End If
    
    ReDim newArr(UBound(arr) - 1)
    For cnt = LBound(newArr) To UBound(newArr)
        newArr(cnt) = arr(cnt)
    Next cnt
    
    Debug.Print printArray(newArr)
    SumArrayRecursion = SumArrayRecursion(newArr) + newArr(UBound(newArr))
    
End Function

Public Function printArray(arr As Variant) As String
    Dim cnt As Long
    For cnt = LBound(arr) To UBound(arr)
        printArray = printArray & " " & arr(cnt)
    Next cnt
End Function
