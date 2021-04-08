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
    Debug.Print "Saved as:" & vbCrLf & newName
    
End Sub

Public Sub SaveThisM()
    
'saves foo.4.5.6.xlsb to foo.4.5.7.xlsb
'and moves the old one to root\Arch\Auto

    Dim oldName As String
    oldName = ThisWorkbook.Name
    
    SaveThis
    
    Dim fso As New FileSystemObject
    fso.MoveFile Source:=ThisWorkbook.path & "\" & oldName, Destination:=ThisWorkbook.path & "\Arch\Auto\" & oldName

    Debug.Print "Moved to:" & vbCrLf & ThisWorkbook.path & "\Arch\Auto\" & oldName
    
End Sub
