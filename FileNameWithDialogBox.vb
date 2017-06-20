Option Explicit

Public Sub MainBrowse(my_obj As Object)
    
    Dim str_file                As String
    
    str_file = Application.GetOpenFilename(Title:="Please choose a file to open", FileFilter:="Excel Files *.xls* (*.xls*),")
    my_obj = str_file

End Sub

Private Sub btnBrowse_Click()
    
    Dim strInitial      As String
    Dim objLabel        As Object
    
    Set objLabel = ThisWorkbook.Worksheets(tbl_input.Name).lblDisplay
    
    strInitial = objLabel
    Call MainBrowse(objLabel)

    If Len(objLabel) >= 6 Then 'Falsch, False
        objLabel = strInitial
    End If

End Sub
