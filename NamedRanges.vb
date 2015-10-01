
Sub change_all_names()
    
    Dim i               As Integer
    Dim s               As String
    Dim s_old           As String
    Dim s_new           As String
    
    For i = 1 To ActiveWorkbook.Names.Count
'        Debug.Print ActiveWorkbook.Names(i).name
'        Debug.Print ActiveWorkbook.Names(i).RefersToR1C1
'        Debug.Print ActiveWorkbook.Names(i)

        If InStr(1, ActiveWorkbook.Names(i), "old", vbTextCompare) Then
            s_old = ActiveWorkbook.Names(i).RefersToR1C1
            s_new = Replace(s_old, "old", "")
            Debug.Print s_new
            
            With ActiveWorkbook.Names(ActiveWorkbook.Names(i).name)
                .RefersToR1C1 = s_new

            End With
        End If
    Next i

End Sub
