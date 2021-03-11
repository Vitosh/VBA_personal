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

Public Function calculate_range(from_row As Long, to_row As Long, l_column As Long, _
                                Optional s_sheet_name As String = "calendar") As Double

    Dim ws              As Worksheet
    Dim l_counter       As Long
    Dim d_result        As Double
    
    Set ws = ThisWorkbook.Worksheets(s_sheet_name)
    
    For l_counter = from_row To to_row
        Call Increment(d_result, ws.Cells(l_counter, l_column))
    Next l_counter

    Set ws = Nothing
    
    calculate_range = Round(d_result, 2)
    
End Function
