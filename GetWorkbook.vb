Public Function GetWorkbook(ByVal sFullName As String) As Workbook
    
    Dim sFile As String
    Dim wbReturn As Workbook
    
    sFile = Dir(sFullName)
    
    On Error Resume Next
        Workbooks(sFile).Close
        Set wbReturn = Workbooks(sFile)
    
        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName)
        End If
    On Error GoTo 0
    
    Set GetWorkbook = wbReturn
    
End Function
