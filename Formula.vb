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
