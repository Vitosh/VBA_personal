' Information for opened Excel Files
' Other Excel Files information
' Opened Excel files
' Excel files count
' Excel count

Public Sub InfoForExcel()

    Dim objList                 As Object
    Dim strProcessName          As String
    
    strProcessName = "EXCEL.EXE"
    
    Set objList = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & strProcessName & "'")
    
    If objList.Count > 1 Then
        MsgBox "Sie haben " & objList.Count & " eröffneten Excel Dateien." & vbCrLf & _
               "Bitte schließen Sie alles, außer der aktuellen Anwendung."
    End If
    
End Sub
