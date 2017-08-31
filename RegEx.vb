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
