Option Explicit

Sub TestMe()
    
    Debug.Print get_relative("U:\DB_DATA\HISTORY_LOG.xlsx")
    Debug.Print get_relative("U:\DB_DATA\HISTORY_LOG.xlsx", 2)

End Sub

Public Function get_relative(str_path As String, Optional l_number As Long = 1) As String

    Dim str_result      As String
    Dim l_start         As Long
    Dim l_counter       As Long
    
    For l_counter = 1 To l_number
        l_start = InStr(l_start + 1, str_path, "\")
    Next l_counter

    get_relative = Mid(str_path, InStr(l_start, str_path, "\"))

End Function
