Option Explicit

Sub GenerateData()
     
    Dim conn            As New ADODB.Connection
    Dim l_row           As Long
    Dim s_username      As String
    Dim s_date          As String
    Dim s_time          As String
    Dim s_location      As String
    Dim s_status        As String
  
    With ActiveSheet
        conn.Open "Provider=SQLOLEDB;Data Source=GRO-PC;Initial Catalog=LogData;Integrated Security=SSPI;"
        
        l_row = last_row_with_data(1, ActiveSheet) + 1
        
        .Cells(l_row, 1) = Environ("username")
        .Cells(l_row, 2) = Date
        .Cells(l_row, 3) = Time
        .Cells(l_row, 4) = Application.ActiveWorkbook.FullName
        .Cells(l_row, 5) = make_random(2, 6)
        
        s_username = .Cells(l_row, 1)
        s_date = .Cells(l_row, 2)
        s_time = .Cells(l_row, 3)
        s_location = .Cells(l_row, 4)
        s_status = .Cells(l_row, 5)
                        
        conn.Execute "insert into dbo.LogTable (UserName, CurrentDate, CurrentTime, CurrentLocation, Status) values ('" & s_username & "', '" & s_date & "', '" & s_time & "', '" & s_location & "','" & s_status & "')"
            
        conn.Close
        Set conn = Nothing

    End With
End Sub

Public Function last_row_with_data(ByVal lng_column_number As Long, shCurrent As Variant) As Long
    last_row_with_data = shCurrent.Cells(Rows.Count, lng_column_number).End(xlUp).row
End Function

Public Function make_random(down As Integer, up As Integer)
    make_random = Int((up - down + 1) * Rnd + down)
End Function

