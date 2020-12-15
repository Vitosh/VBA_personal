Public Function b_file_exists(ByVal str_file_path As String) As Boolean

    Dim str_test    As String
    
    On Error Resume Next
    str_test = Dir(str_file_path)
    On Error GoTo 0
    b_file_exists = (str_test <> "")

End Function

'works in eshare
'eshare file exists

Public Function EshareFileExists(filePath)
    
    filePath = Replace(filePath, "https:", "")
    filePath = Replace(filePath, "%20", " ")
    filePath = Replace(filePath, "/", "\")
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
    
End Function
