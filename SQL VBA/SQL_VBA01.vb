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
    
    b_insert_into_table = True

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
