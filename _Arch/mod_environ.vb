Option Explicit

Declare Function GetLocaleInfo Lib "kernel32" Alias _
"GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
ByVal lpLCData As String, ByVal cchData As Long) As Long

Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Const LOCALE_SLIST = &HC
Public Function GetListSeparator() As String

    '?environ("pathext")

    Dim ListSeparator       As String
    Dim iRetVal1            As Long
    Dim iRetVal2            As Long
    Dim lpLCDataVar         As String
    
    Dim Position            As Integer
    Dim Locale              As Long
    
    Locale = GetUserDefaultLCID()
    
    iRetVal1 = GetLocaleInfo(Locale, LOCALE_SLIST, lpLCDataVar, 0)
    
    ListSeparator = String$(iRetVal1, 0)
    
    iRetVal2 = GetLocaleInfo(Locale, LOCALE_SLIST, ListSeparator, iRetVal1)
    
    Position = InStr(ListSeparator, Chr$(0))
    
    If Position > 0 Then
        GetListSeparator = Left$(ListSeparator, Position - 1)
    End If

End Function

Sub EnumSEVars()
    Dim strVar As String
    Dim i As Long
    
    For i = 1 To 255
        strVar = Environ$(i)
        If LenB(strVar) = 0& Then Exit For
        Debug.Print strVar
    Next
End Sub

