Public Sub PrintMeUsefulFormula()

    Dim strFormula  As String
    Dim strParenth  As String

    strParenth = """"

    strFormula = Selection.Formula
    strFormula = Replace(strFormula, """", """""")

    strFormula = strParenth & strFormula & strParenth
    Debug.Print strFormula

End Sub
