Option Explicit

'---------------------------------------------------------------------------------------
' Method : CompareVersions
' Author : v.doynov
' Date   : 08.12.2016
' Purpose: Two public subs - PostInfo and CompareVersions
'---------------------------------------------------------------------------------------

Private version_sql     As String
Private date_sql        As Date

Public Function CompareVersions() As Boolean

    If (Me.DateSQL = Me.DateWorkbook) And (Me.VersionSQL = Me.VersionWorkbook) Then
        CompareVersions = True
    Else
        CompareVersions = False
    End If

End Function

Private Function str_connection_string() As String

    Dim arr_info(5) As Variant

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

Private Function str_generator(ByVal str_value As String, ByVal b_fix As Boolean) As String

    Dim l_counter As Long
    Dim l_number As Long
    Dim str_char As String

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

Private Function b_is_odd(l_number As Long) As Boolean

    b_is_odd = l_number Mod 2

End Function

Public Property Get VersionWorkbook() As String

    VersionWorkbook = [set_version_number]

End Property

Public Property Get DateWorkbook() As Date

    DateWorkbook = [set_version_date]

End Property

Public Property Get VersionSQL() As String

    VersionSQL = version_sql

End Property

Public Property Get DateSQL() As Date

    DateSQL = date_sql

End Property

Public Function str_post_info() As String

    str_post_info = "  Diese Version ist - " & Me.VersionWorkbook & " von " & Me.DateWorkbook & "." & vbCrLf & _
                  "  Die letzte ist          - " & Me.VersionSQL & " von " & Me.DateSQL & "."

End Function

Public Sub GetDataFromSQLServer()

    If [set_in_production] Then On Error GoTo GetDataFromSQLServer_Error

    Dim cnLogs As Object
    Dim rsData As Object

    Set cnLogs = CreateObject("ADODB.Connection")
    Set rsData = CreateObject("ADODB.Recordset")

    cnLogs.Open str_connection_string
    cnLogs.Execute "SET NOCOUNT ON"

    With rsData
        .ActiveConnection = cnLogs
        .Open "SELECT [VersionNumber],[MyDate] FROM [Versions] WHERE IsLastCurrent=1;"
        version_sql = rsData.Fields("VersionNumber").value
        date_sql = rsData.Fields("MyDate").value
    End With

    rsData.Close
    cnLogs.Close

    Set cnLogs = Nothing
    Set rsData = Nothing

    On Error GoTo 0
    Exit Sub

GetDataFromSQLServer_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDataFromSQLServer of Sub cls_Version"
    Set cnLogs = Nothing
    Set rsData = Nothing
    version_sql = [set_version_check_error]
    date_sql = [set_version_check_error]

End Sub
