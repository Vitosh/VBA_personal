Option Explicit

Public Sub GenerateDataIntoTable()

    Dim str_table_name      As String: str_table_name = "Main"
    Dim arr_column_names    As Variant
    Dim arr_values          As Variant
    
    ReDim arr_column_names(6)
    ReDim arr_values(6)
    
    arr_column_names(0) = "UserName"
    arr_column_names(1) = "CurrentDate"
    arr_column_names(2) = "CurrentTime"
    arr_column_names(3) = "CurrentLocation"
    arr_column_names(4) = "Status1"
    arr_column_names(5) = "Status2"
    arr_column_names(6) = "Status3"
    
    arr_values(0) = Environ("username")
    arr_values(1) = Date
    arr_values(2) = Time
    arr_values(3) = Application.ActiveWorkbook.FullName
    arr_values(4) = make_random(2, 6)
    arr_values(5) = arr_values(4) + make_random(2, 6)
    arr_values(6) = arr_values(5) - make_random(2, 6)

    Debug.Print b_insert_into_table(str_table_name, arr_column_names, arr_values)

End Sub

Function b_insert_into_table(str_table_name As String, arr_column_names As Variant, arr_values As Variant) As Boolean

    Dim conn            As Object
    Dim str_order       As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open str_connection_string
    
    str_order = "insert into dbo." & str_table_name
    str_order = str_order & str_generate_order(arr_column_names, arr_values)
    conn.Execute str_order
    conn.Close
    Set conn = Nothing

End Function

Public Function str_generate_order(arr_column_names As Variant, arr_values As Variant) As String

    Dim l_counter       As Long
    Dim str_result      As String
    
    Dim str_left        As String: str_left = "('"
    Dim str_midd        As String: str_midd = "','"
    Dim str_right       As String: str_right = "')"
    
    str_result = "("
    For l_counter = LBound(arr_column_names) To UBound(arr_column_names)
        str_result = str_result & arr_column_names(l_counter) & ","
    Next l_counter
    
    str_result = Left(str_result, Len(str_result) - 1)
    str_result = str_result & ")"
    str_result = str_result & "values"
    
    str_result = str_result & str_left
    For l_counter = LBound(arr_values) To UBound(arr_values)
        str_result = str_result & arr_values(l_counter)
        
        If l_counter < UBound(arr_values) Then
            str_result = str_result & str_midd
        Else
            str_result = str_result & str_right
        End If
        
    Next l_counter
    
    str_generate_order = str_result
    
End Function

Sub GenerateData()

    Dim conn            As Object
    Dim l_row           As Long

    Dim s_username      As String
    Dim s_date          As String
    Dim s_time          As String
    Dim s_location      As String
    Dim s_status        As String

    Set conn = CreateObject("ADODB.Connection")
    
    With ActiveSheet
        conn.Open str_connection_string
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
    End With

    conn.Execute "insert into dbo.Main (UserName, CurrentDate, CurrentTime, CurrentLocation, Status1, Status2, Status3) values ('" & s_username & "', '" & s_date & "', '" & s_time & "', '" & s_location & "','" & s_status & "','" & s_status + 2 & "','" & s_status + 3 & "')"
    conn.Close
    Set conn = Nothing

End Sub

Sub GetData()

    Dim cnLogs              As Object
    Dim rsHeaders           As Object
    Dim rsData              As Object
    
    Dim l_counter           As Long: l_counter = 0
    Dim strConn             As String
    
    Set cnLogs = CreateObject("ADODB.Connection")
    Set rsHeaders = CreateObject("ADODB.Recordset")
    Set rsData = CreateObject("ADODB.Recordset")
    
    Sheets(1).UsedRange.Clear
    cnLogs.Open str_connection_string
    
    With rsHeaders
        .ActiveConnection = cnLogs
        
        .Open "SELECT * FROM syscolumns WHERE id=OBJECT_ID('Main')"
        
        Do While Not rsHeaders.EOF
            Cells(1, l_counter + 1) = rsHeaders(0)
            l_counter = l_counter + 1
            rsHeaders.MoveNext
        Loop
        .Close
    End With

    With rsData
        .ActiveConnection = cnLogs
        .Open "SELECT * FROM Main"
        Sheets(1).Range("A2").CopyFromRecordset rsData
        .Close
    End With
    
    cnLogs.Close
    Set cnLogs = Nothing
    Set rsHeaders = Nothing
    Set rsData = Nothing
    
    Sheets(1).UsedRange.EntireColumn.AutoFit


End Sub

Public Function str_connection_string() As String
    
    Dim arr_info(5)     As Variant
    
    arr_info(0) = [set_conn_provider]
    arr_info(1) = [set_conn_data_source]
    arr_info(2) = [set_conn_database]
    arr_info(3) = [set_conn_user_id]
    arr_info(4) = [set_conn_password]
    
    str_connection_string = "Provider=" & arr_info(0) & _
                    "; Data Source=" & arr_info(1) & _
                    "; Database=" & arr_info(2) & _
                    ";User ID=" & str_generator(arr_info(3), True) & _
                    "; Password=" & str_generator(arr_info(4), True) & ";"
End Function

Public Function last_row_with_data(ByVal lng_column_number As Long, shCurrent As Variant) As Long

    last_row_with_data = shCurrent.Cells(Rows.Count, lng_column_number).End(xlUp).Row
    
End Function
 
Public Function make_random(down As Long, up As Long)

    make_random = Int((up - down + 1) * Rnd + down)
    
End Function

Public Function str_generator(ByVal str_value As String, ByVal b_fix As Boolean) As String
    
    Dim l_counter   As Long
    Dim l_number    As Long
    Dim str_char    As String
    
    On Error GoTo str_generator_Error
    
    If b_fix Then
        str_value = Left(str_value, Len(str_value) - 1)
        str_value = Right(str_value, Len(str_value) - 1)
    End If

    For l_counter = 1 To Len(str_value)
        str_char = Mid(str_value, l_counter, 1)
        If b_is_odd(l_counter) Then
            l_number = Asc(str_char) + IIf(b_fix, -2, 2)
        Else
            l_number = Asc(str_char) + IIf(b_fix, -3, 3)
        End If
        
        str_generator = str_generator + Chr(l_number)
    
    Next l_counter
    
    If Not b_fix Then
        str_generator = Chr(l_number) & str_generator & Chr(l_number)
    End If
    
    On Error GoTo 0
    Exit Function

str_generator_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure str_generator of Function Modul1"

End Function

Public Function b_is_odd(l_number As Long) As Boolean

    b_is_odd = l_number Mod 2
    
End Function