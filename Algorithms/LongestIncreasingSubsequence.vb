Option Explicit

Public Const NO_PREVIOUS = -1

Sub Main()

    Dim arrSeq         As Variant
    Dim arrLen         As Variant
    Dim arrPre         As Variant
    
    Dim bestLength        As Long
    
    arrSeq = Array(1, 2, -6, -5, -3, 23, 123, 3, 2, -23, -5, 54, 100, 200, 300, 1111, 23412, 3, 4, 5, 6, 7, 8, 9, 19, 65, 2)
    ReDim arrLen(UBound(arrSeq))
    ReDim arrPre(UBound(arrSeq))
    
    bestLength = CalculateLongestIncreasingSubsequence(arrSeq, arrLen, arrPre)
    PrintArray arrSeq
    PrintArray arrLen
    PrintArray arrPre
    
    PrintLongestIncreasingSubsequance arrSeq, arrPre, bestLength
    
End Sub

Public Sub PrintLongestIncreasingSubsequance(ByRef arrSeq As Variant, _
                                            ByRef arrPre As Variant, _
                                            bestLength As Long)
                                            
    Dim arrResult  As Variant
    Dim counter As Long: counter = 0
    
    ReDim arrResult(1)
    
    While (bestLength <> NO_PREVIOUS)
        ReDim Preserve arrResult(counter)
        counter = counter + 1
        arrResult(counter - 1) = arrSeq(bestLength)
        bestLength = arrPre(bestLength)
    Wend
    
    Debug.Print Join(ReverseArray(arrResult), " ")
    
End Sub


Public Function CalculateLongestIncreasingSubsequence(ByRef arrSeq As Variant, _
                                                    ByRef arrLen As Variant, _
                                                    ByRef arrPre As Variant) As Long

    Dim bestLengthLen    As Long: bestLengthLen = 0
    Dim bestLengthIndex    As Long: bestLengthIndex = 0
    Dim x               As Long
    Dim i               As Long
    
    For x = LBound(arrSeq) To (UBound(arrSeq))
        arrLen(x) = 1
        arrPre(x) = NO_PREVIOUS
        
        For i = 0 To x Step 1
            If (arrSeq(i) < arrSeq(x)) And (arrLen(i) + 1 > arrLen(x)) Then
                
                arrLen(x) = arrLen(i) + 1
                arrPre(x) = i
                
                If arrLen(x) > bestLengthLen Then
                    bestLengthLen = arrLen(x)
                    bestLengthIndex = x
                End If
            End If
            
        Next i
    Next x
        
    CalculateLongestIncreasingSubsequence = bestLengthIndex
    
End Function

Public Sub PrintArray(ByRef myArray As Variant)
    Dim counter As Long
    
    For counter = LBound(myArray) To UBound(myArray)
        Debug.Print counter & " --> " & myArray(counter)
    Next counter
    Debug.Print "------------------------------"
End Sub

Public Function ReverseArray(ByVal myArray As Variant) As Variant

    Dim counter     As Long
    Dim counter2   As Long
    Dim arrNew     As Variant
    
    ReDim arrNew(UBound(myArray) + 1)
    
    For counter = LBound(arrNew) To UBound(arrNew) - 1
        counter2 = UBound(arrNew) - counter - 1
        arrNew(counter) = myArray(counter2)
    Next counter

    ReverseArray = arrNew

End Function

