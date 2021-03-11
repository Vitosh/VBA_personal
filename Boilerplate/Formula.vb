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

'Column to letter letter to column
'lettertocolumn columntoletter

Function ColumnToLetter(columnNumber As Long) As String
   
    If columnNumber < 1 Then Exit Function
    ColumnToLetter = UCase(ColumnToLetter(Int((columnNumber - 1) / 26)) & Chr(((columnNumber - 1) Mod 26) + Asc("A")))

End Function

Function LetterToColumn(letters As String) As Long
    
    Dim i As Long
    letters = UCase(letters)
    
    For i = Len(letters) To 1 Step -1
        LetterToColumn = LetterToColumn + (Asc(Mid(letters, i, 1)) - 64) * 26 ^ (Len(letters) - i)
    Next
        
End Function

Sub Tests()

    Debug.Print LetterToColumn("a") = 1
    Debug.Print LetterToColumn("A") = 1
    Debug.Print LetterToColumn("Z") = 26
    Debug.Print LetterToColumn("AA") = 27
    Debug.Print LetterToColumn("AZ") = 52
    Debug.Print LetterToColumn("BA") = 53
    
    Debug.Print ColumnToLetter(1) = "A"
    Debug.Print ColumnToLetter(26) = "Z"
    Debug.Print ColumnToLetter(27) = "AA"
    Debug.Print ColumnToLetter(52) = "AZ"
    Debug.Print ColumnToLetter(53) = "BA"
    
End Sub
