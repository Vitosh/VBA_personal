Public Sub PrintMeUsefulFormula()

    Dim selectedFormula  As String
    Dim parenthesis  As String

    parenthesis = """"

    selectedFormula = Selection.Formula
    selectedFormula = Replace(selectedFormula, """", """""")

    selectedFormula = parenthesis & selectedFormula & parenthesis
    Debug.Print selectedFormula
    
End Sub

'A bit untested, use with caution --------v
Public Sub PrintMeUsefulFormat()

    Dim strFormula  As String
    Dim strParenth  As String

    strParenth = """"

    strFormula = Selection.NumberFormat
    strFormula = Replace(strFormula, """", """""")

    strFormula = strParenth & strFormula & strParenth
    Debug.Print strFormula

End Sub

'Column to letter lettertocolumn columntoletter letter to column

Function ColumnToLetter(columnNumber As Long) As String
   
    If columnNumber < 1 Then Exit Function
    ColumnToLetter = ColumnToLetter(Int((columnNumber - 1) / 26)) & Chr(((columnNumber - 1) Mod 26) + asc("A"))

End Function
