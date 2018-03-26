Option Explicit

Sub TestMe()

    Dim myArr           As Variant
    Dim myLoop          As Variant
    Dim targetValue     As Long
    Dim currentSum      As Long

    myArr = Array(215, 275, 335, 355, 420, 580)
    targetValue = 1505

    Dim cnt0&, cnt1&, cnt2&, cnt3&, cnt4&, cnt5&, cnt6&
    Dim cnt As Long


    For cnt0 = 0 To 5
        For cnt1 = 0 To 5
            For cnt2 = 0 To 5
                For cnt3 = 0 To 5
                    For cnt4 = 0 To 5
                        For cnt5 = 0 To 5
                            currentSum = 0

                            Dim printableArray As Variant
                            printableArray = Array(cnt0, cnt1, cnt2, cnt3, cnt4, cnt5)

                            For cnt = LBound(myArr) To UBound(myArr)
                                IncrementSum printableArray(cnt), myArr(cnt), currentSum
                            Next cnt

                            If currentSum = targetValue Then
                                printValuesOfArray printableArray, myArr
                            End If
    Next: Next: Next: Next: Next: Next

End Sub

Public Sub printValuesOfArray(myArr As Variant, initialArr As Variant)

    Dim cnt             As Long
    Dim printVal        As String

    For cnt = LBound(myArr) To UBound(myArr)
        If myArr(cnt) Then
            printVal = printVal & myArr(cnt) & " * " & initialArr(cnt) & vbCrLf
        End If
    Next cnt

    Debug.Print printVal

End Sub

Public Sub IncrementSum(ByVal multiplicator As Long, _
    ByVal arrVal As Long, ByRef currentSum As Long)

    currentSum = currentSum + arrVal * multiplicator

End Sub
