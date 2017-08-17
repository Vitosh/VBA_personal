Public Sub PrintMeUsefulFormula()

    Dim strFormula  As String
    Dim strParenth  As String

    strParenth = """"

    strFormula = Selection.Formula
    strFormula = Replace(strFormula, """", """""")

    strFormula = strParenth & strFormula & strParenth
    Debug.Print strFormula

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
