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
    EshareFileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
    
End Function

Public Sub SaveThis()
    
    'saves foo.4.5.6.xlsb to foo.4.5.7.xlsb
    
    Dim mySplitter As Variant
    mySplitter = Split(ThisWorkbook.FullName, ".")
    
    Dim oldVersion As String
    oldVersion = mySplitter(UBound(mySplitter) - 1)
    
    Dim newVersion As String
    newVersion = oldVersion + 1
    
    mySplitter(UBound(mySplitter) - 1) = newVersion
    
    Dim newName As String
    newName = Join(mySplitter, ".")
    
    ThisWorkbook.SaveAs newName
    Debug.Print newName
    
End Sub
