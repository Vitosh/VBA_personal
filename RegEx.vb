Option Explicit

Public Sub RegExExample()
    
    Dim strString       As String
    Dim lngCounter      As Long
    Dim objRegex        As Object
    Dim arrWords        As Variant
    
    'RegEx with late binding
    Set objRegex = CreateObject("VBScript.RegExp")

    strString = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."
    arrWords = Split(strString)
    objRegex.Pattern = "or"
    
    For lngCounter = LBound(arrWords) To UBound(arrWords)
        If objRegex.test(arrWords(lngCounter)) Then
            Debug.Print arrWords(lngCounter)
        End If
    Next lngCounter

End Sub

'===============================================================================
'===============================================================================
'removes anything that is not a digit or word from the string===================

Public Function removeInvisibleThings(s As String) As String

    Dim regEx           As Object
    Dim inputMatches    As Object
    Dim regExString     As String

    Set regEx = CreateObject("VBScript.RegExp")

    With regEx
        .pattern = "[^a-zA-Z0-9]"
        .IgnoreCase = True
        .Global = True

        Set inputMatches = .Execute(s)

        If regEx.test(s) Then
            removeInvisibleThings = .Replace(s, vbNullString)
        Else
            removeInvisibleThings = s
        End If

    End With

End Function

Public Sub TestMe()

    Debug.Print removeInvisibleThings("aa1 Abc 67 ( *^ 45 ")
    Debug.Print removeInvisibleThings("aa1 ???!")
    Debug.Print removeInvisibleThings("   aa1 Abc 1267 ( *^ 45 ")

End Sub

'===============================================================================
'===============================================================================
'===============================================================================

Public Function findTheSubString(wholeString As String, subString As String) As String

    Dim regEx           As Object
    Dim inputMatches    As Object
    Dim regExString     As String

    Set regEx = CreateObject("VBScript.RegExp")

    With regEx
        .Pattern = Split(subString, "*")(0) & "[\s\S]*" & Split(subString, "*")(1)
        .IgnoreCase = True
        .Global = True

        Set inputMatches = .Execute(wholeString)
        If regEx.test(wholeString) Then
            findTheSubString = inputMatches(0)
        Else
            findTheSubString = "Not Found!"
        End If

    End With

End Function

'===============================================================================
'===============================================================================
'===============================================================================
