Option Explicit

Public Sub SplitSingleColumnToCells()

    Dim rngInput    As Range
    Dim rngOutput   As Range
    Dim myCell      As Range

    'Set manually, it is faster :)
    Set rngInput = Range("A1:A22")

    For Each myCell In rngInput
        'replace multiple space with single space:
        myCell = Replace(myCell, Chr(32), Chr(32))
        Dim inputArray As Variant
        inputArray = Split(myCell)

        Dim col     As Long
        Dim i       As Long
        col = 0
        For i = LBound(inputArray) To UBound(inputArray)
            If Len(inputArray(i)) > 0 Then
                col = col + 1
                myCell.Offset(0, col) = inputArray(i)
            End If
        Next i
        'Probably not needed:
        'myCell.Clear
    Next myCell
End Sub
