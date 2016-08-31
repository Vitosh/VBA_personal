Option Explicit

Sub ServerUpload()

    Dim conn            As Object
    Dim l_last_row      As Long
    
    Dim l_counter       As Long
    Dim l_counter2      As Long
    
    Dim str_left        As String
    Dim str_right       As String
    
    If Application.WorksheetFunction.CountIf(tbl_summary.UsedRange, ERROR_NUMBER) > 0 Then
        MsgBox "Keine roten Werte erlaubt!", vbInformation, "TEMPTM"
        Exit Sub
    End If
    
    Set conn = CreateObject("ADODB.Connection")
    l_last_row = last_row(tbl_summary.Name)

    For l_counter = 2 To l_last_row Step 1
        conn.Open str_connection_string
        
        str_right = "('" & Date & "','" & Time & "','" & Environ("Username") & "','" & tbl_summary.Cells(l_counter, 2) & "',"
        
        For l_counter2 = 3 To 17 Step 1
            str_right = str_right & Str(tbl_summary.Cells(l_counter, l_counter2)) & ","
        Next l_counter2
        
        str_right = Left(str_right, Len(str_right) - 1) & ")"
        str_left = "(~,~,~,~,~,~,~,"
        str_left = str_left & ~,~,~,~,~,~,~)"
        
        conn.Execute "insert into dbo.tempt_report" & str_left & "VALUES" & str_right
        conn.Close
    Next l_counter
         
    Set conn = Nothing
    Debug.Print "UPLOAD SUCCESSFUL!"
    
End Sub

Sub ResetInfoInTable()

    Dim cnLogs              As Object

    If Not b_value_in_array(str_get_username, ADMINS, True) Then Exit Sub
    
    Set cnLogs = CreateObject("ADODB.Connection")

    cnLogs.Open str_connection_string
    cnLogs.Execute "TRUNCATE TABLE tempt_report;"
    cnLogs.Close
    Set cnLogs = Nothing

    Debug.Print "TABLE tempt_report has been truncated"

End Sub

Public Function str_get_username() As String
    
    str_get_username = Environ("Username")
    
End Function

Sub ServerDownload()
    
    Dim cnLogs              As Object
    Dim rsHeaders           As Object
    Dim rsData              As Object
    
    Dim l_counter           As Long

    Call OnStart
    
    If Not b_value_in_array(str_get_username, ADMINS, True) Then Exit Sub
    
    Set cnLogs = CreateObject("ADODB.Connection")
    Set rsHeaders = CreateObject("ADODB.Recordset")
    Set rsData = CreateObject("ADODB.Recordset")
    
    tbl_all.UsedRange.Clear
    cnLogs.Open str_connection_string
    
    With rsHeaders
        .ActiveConnection = cnLogs
        .Open "SELECT * FROM syscolumns WHERE id=OBJECT_ID('tempt_report')"
        
        Do While Not rsHeaders.EOF
            Cells(1, l_counter + 1) = rsHeaders(0)
            l_counter = l_counter + 1
            rsHeaders.MoveNext
        Loop
        .Close
    End With
    
    With rsData
        .ActiveConnection = cnLogs
        .Open "SELECT * FROM tempt_report;"
        tbl_all.Cells(2, 1).CopyFromRecordset rsData
        .Close
    End With
    
    Call FormatCells
    Call OnEnd
    
    Debug.Print "DOWNLOAD SUCCESSFUL!"
    
End Sub

Sub FormatCells()

    Dim l_rows              As Long
    Dim l_cols              As Long
    
    Dim l_counter           As Long
    Dim l_counter2          As Long
    
    Dim my_cell             As Range
    
    Call OnStart
    
    l_cols = last_column(tbl_all.Name)
    l_rows = last_row(tbl_all.Name)
    
    For l_counter = 1 To l_cols
        For l_counter2 = 2 To l_rows
            
            Set my_cell = tbl_all.Cells(l_counter2, l_counter)
            
            Select Case True
                Case tbl_all.Cells(1, l_counter) = "Datum"
                    my_cell.NumberFormat = "[$-407]d/ mmm/ yy;@"
                    my_cell.FormulaR1C1 = my_cell.Text
                Case tbl_all.Cells(1, l_counter) = "Zeit"
                    
                    my_cell.FormulaR1C1 = Split(my_cell, ".")(0)
                        my_cell.NumberFormat = "hh:mm"
                    
                Case tbl_all.Cells(1, l_counter) = "IRR" Or tbl_all.Cells(1, l_counter) = "ObjektReturn"
                    my_cell.NumberFormat = "0.00%"
                Case tbl_all.Cells(1, l_counter) <> "ID" _
                            And tbl_all.Cells(1, l_counter) <> "Benutzer" _
                            And tbl_all.Cells(1, l_counter) <> "Objekt"
                    my_cell.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            End Select
        Next l_counter2
    Next l_counter
        
    Set my_cell = Nothing
    Call OnEnd

End Sub
