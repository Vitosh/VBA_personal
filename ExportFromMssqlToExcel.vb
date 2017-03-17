Option Explicit

'The part extracting the body is taken from here
'https://support.microsoft.com/en-us/kb/306125

Sub GetData()
    
    Dim cnLogs              As New ADODB.Connection
    Dim rsHeaders           As New ADODB.Recordset
    Dim rsData              As New ADODB.Recordset
    
    Dim l_counter           As Long: l_counter = 0
    Dim strConn             As String
    
    Sheets(1).UsedRange.Clear
    strConn = "PROVIDER=SQLOLEDB;"
    strConn = strConn & "DATA SOURCE=(local);INITIAL CATALOG=LogData;"
    strConn = strConn & " INTEGRATED SECURITY=sspi;"
    
    cnLogs.Open strConn
    
    With rsHeaders
        .ActiveConnection = cnLogs
        
        .Open "SELECT * FROM syscolumns WHERE id=OBJECT_ID('LogTable')"
        '.Open "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'LogTable'"
        '.Open "SELECT * FROM LogData.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = N'LogTable'"
        '.Open "SELECT * FROM SYS.COLUMNS WHERE object_id = OBJECT_ID('dbo.LogTable')"
        
        Do While Not rsHeaders.EOF
            Cells(1, l_counter + 1) = rsHeaders(0)
            l_counter = l_counter + 1
            rsHeaders.MoveNext
        Loop
        .Close
    End With

    With rsData
        .ActiveConnection = cnLogs
        .Open "SELECT * FROM LogTable"
        Sheet1.Range("A2").CopyFromRecordset rsData
        .Close
    End With
    
    cnLogs.Close
    Set cnLogs = Nothing
    Set rsHeaders = Nothing
    Set rsData = Nothing
    
    Sheets(1).UsedRange.EntireColumn.AutoFit

End Sub

