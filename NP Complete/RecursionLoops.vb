Option Explicit

Sub Main()

    Dim posArr                  As Variant
    Dim iniArr                  As Variant
    Dim tryArr                  As Variant
    Dim cnt                     As Long
    Dim targetVal               As Long: targetVal = 1505

    iniArr = Array(215, 275, 335, 355, 420, 580)
    ReDim posArr(UBound(iniArr))
    ReDim tryArr(UBound(iniArr))

    For cnt = LBound(posArr) To UBound(posArr)
        posArr(cnt) = cnt
    Next cnt
    EmbeddedLoops 0, posArr, tryArr, iniArr, targetVal

End Sub

Function EmbeddedLoops(index As Long, posArr As Variant, tryArr As Variant, _
                                      iniArr As Variant, targetVal As Long)

    Dim myUnit              As Variant
    Dim cnt                 As Long

    If index >= UBound(posArr) + 1 Then
        If CheckSum(tryArr, iniArr, targetVal) Then
            For cnt = LBound(tryArr) To UBound(tryArr)
                If tryArr(cnt) Then Debug.Print tryArr(cnt) & " x " & iniArr(cnt)
            Next cnt
        End If
    Else
        For Each myUnit In posArr
            tryArr(index) = myUnit
            EmbeddedLoops index + 1, posArr, tryArr, iniArr, targetVal
        Next myUnit
    End If

End Function

Public Function CheckSum(posArr, iniArr, targetVal) As Boolean

    Dim cnt         As Long
    Dim compareVal  As Long

    For cnt = LBound(posArr) To UBound(posArr)
        compareVal = posArr(cnt) * iniArr(cnt) + compareVal
    Next cnt
    CheckSum = CBool(compareVal = targetVal)

End Function
