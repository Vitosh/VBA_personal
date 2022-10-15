Function InsertIntoString(originalString As String, addedString As String, positionToAdd As Long) As String

    If positionToAdd < 1 Then positionToAdd = 1
    If Len(originalString) < positionToAdd Then positionToAdd = Len(originalString) + 1

    InsertIntoString = Mid(originalString, 1, positionToAdd - 1) _
                        & addedString _
                        & Mid(originalString, positionToAdd, Len(originalString) - positionToAdd + 1)

End Function

Public Sub TestInsertIntoString()

    Debug.Print InsertIntoString("abcd", "ff", 0) = "ffabcd"
    Debug.Print InsertIntoString("abcd", "ff", 1) = "ffabcd"
    Debug.Print InsertIntoString("abcd", "ff", 2) = "affbcd"
    Debug.Print InsertIntoString("abcd", "ff", 3) = "abffcd"
    Debug.Print InsertIntoString("abcd", "ff", 4) = "abcffd"
    Debug.Print InsertIntoString("abcd", "ff", 100) = "abcdff"

End Sub

