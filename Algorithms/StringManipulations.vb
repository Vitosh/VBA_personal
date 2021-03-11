Function Insert(original As String, added As String, pos As Long) As String
    
    If pos < 1 Then pos = 1
    If Len(original) < pos Then pos = Len(original) + 1
    
    Insert = Mid(original, 1, pos - 1) _
                        & added _
                        & Mid(original, pos, Len(original) - pos + 1)
    
End Function

Public Sub InsertTests()

    Debug.Print Insert("abcd", "ff", 0) = "ffabcd"
    Debug.Print Insert("abcd", "ff", 1) = "ffabcd"
    Debug.Print Insert("abcd", "ff", 2) = "affbcd"
    Debug.Print Insert("abcd", "ff", 3) = "abffcd"
    Debug.Print Insert("abcd", "ff", 4) = "abcffd"
    Debug.Print Insert("abcd", "ff", 100) = "abcdff"
    
End Sub

Public Function StringRepeater(repeatString As String, count As Long) As String
    'StringBuilder String Builder 
    If count < 1 Or Len(repeatString) < 1 Then Exit Function
    
    Dim cnt As Long
    
    For cnt = 1 To count
        StringRepeater = StringRepeater & repeatString
    Next cnt

End Function

Public Sub StringRepeaterTests()

    Debug.Print StringRepeater("ab", 3) = "ababab"
    Debug.Print StringRepeater("a", 2) = "aa"
    
End Sub
