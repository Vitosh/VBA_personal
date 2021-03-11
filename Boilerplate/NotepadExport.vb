' export to notepad export txt export string string to txt string to notepad

Option Explicit

Public STR_ERROR_REPORT                 As String

Sub CreateLogFile(Optional str_print As String)

    On Error GoTo CreateLogFile_Error

    Dim fs                      As Object
    Dim obj_text                As Object
    Dim str_filename            As String
    Dim str_new_file            As String
    Dim str_shell               As String

    str_new_file = "\tests_info"

    str_filename = ThisWorkbook.Path & str_new_file & codify_time(True)
    If Dir(ThisWorkbook.Path & str_new_file, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & str_new_file

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set obj_text = fs.CreateTextFile(str_filename, True)

    If Len(STR_ERROR_REPORT) > 1 Then
        obj_text.writeline (STR_ERROR_REPORT)
    Else
        obj_text.writeline (str_print)
    End If
    
    obj_text.Close

    str_shell = "C:\WINDOWS\notepad.exe "
    str_shell = str_shell & str_filename
    Call Shell(str_shell)

    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"
    
End Sub

Public Function codify_time(Optional b_make_str As Boolean = False) As String

    On Error GoTo codify_Error
    
    Dim dbl_01                  As Variant
    Dim dbl_02                  As Variant
    Dim dbl_now                 As Double
    
    dbl_now = Round(Now(), 8)
    
    dbl_01 = Split(CStr(dbl_now), ",")(0)
    dbl_02 = Split(CStr(dbl_now), ",")(1)
    
    codify_time = Hex(dbl_01) & "_" & Hex(dbl_02)
    
    If b_make_str Then codify_time = "\" & codify_time & ".txt"
    
    On Error GoTo 0
    Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function
