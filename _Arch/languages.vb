'change language
'change fonts
Option Explicit

Public Enum LandName
    BG
    US
    DE
End Enum

Private Const LOCALE_ILANGUAGE      As Long = &H1
Private Const LOCALE_SCOUNTRY       As Long = &H6
Private Declare Function ActivateKeyboardLayout Lib "user32.dll" (ByVal myLanguage As Long, Flag As Boolean) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long

Declare Function getUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long


Private Declare Function GetLocaleInfo Lib "kernel32" _
            Alias "GetLocaleInfoA" _
            (ByVal Locale As Long, _
            ByVal LCType As Long, _
            ByVal lpLCData As String, _
            ByVal cchData As Long) As Long
    
Public Function f_str_country_name(l_landname As Long) As String
    
    Dim str_result      As String
    
    Select Case l_landname
    
    Case 0:
        str_result = "Bulgarien"
    Case 1:
        str_result = "Vereinigte Staaten"
    Case 2:
        str_result = "Deutschland"
    End Select
    
    f_str_country_name = str_result
    
End Function

Public Function f_lng_country_code(l_landname As Long) As Long
    
    Dim lng_result          As Long
    
    Select Case l_landname
    
    Case 0:
        lng_result = 1026
    Case 1:
        lng_result = 1033
    Case 2:
        lng_result = 1031
    End Select
    
    f_lng_country_code = lng_result
    
End Function

Public Sub ChangeLanguages()

    Call SetLanguage(f_str_country_name(LandName.DE), f_lng_country_code(LandName.DE))

    Call SetLanguage(f_str_country_name(LandName.BG), f_lng_country_code(LandName.BG))

    Call SetLanguage(f_str_country_name(LandName.US), f_lng_country_code(LandName.US))

    Call SetLanguage
    
End Sub
    
Public Sub SetLanguage(Optional str_lang As String = "Bulgarien", Optional l_code As Long = 1026)
    
    If Not f_str_get_language = str_lang Then
       ActivateKeyboardLayout l_code, 0
    End If
    
End Sub

Public Function f_str_get_language()

    Dim hKeyboardID As Long
    Dim LCID As Long
    
    hKeyboardID = GetKeyboardLayout(0&)
    LCID = LoWord(hKeyboardID)

    f_str_get_language = GetUserLocaleInfo(LCID, LOCALE_SCOUNTRY)

End Function

Private Function LoWord(wParam As Long) As Long

    If wParam And &H8000& Then
        LoWord = &H8000& Or (wParam And &H7FFF&)
    Else
        LoWord = wParam And &HFFFF&
    End If
    
End Function

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

    Dim sReturn     As String
    Dim nSize       As Long
    
    nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
    If nSize > 0 Then
        sReturn = Space$(nSize)
        nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
        If nSize > 0 Then
            GetUserLocaleInfo = Left$(sReturn, nSize - 1)
        End If
    End If
    
End Function
