
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

Public Sub MakeNegativesOne(l_col As Long)

    Dim l_counter           As Long
    Dim b_negative          As Long
    Dim my_cell             As Range
    Dim my_first_negative   As Range
    
    Dim dbl_negative_sum    As Double
    
    For l_counter = 1 To 13
        Set my_cell = Cells(l_col, l_counter)
        
        If my_cell < 0 And my_cell.HasFormula Then
            dbl_negative_sum = dbl_negative_sum + my_cell.Value
            
            If Not b_negative Then
                b_negative = True
                Set my_first_negative = my_cell
            End If
            
            my_cell = 0
        End If
    Next l_counter
    
    If b_negative Then
        my_first_negative = dbl_negative_sum
    End If
    
End Sub

Public Sub NegativeSelection(Optional my_rng As Variant)

    Dim my_cell As Range

    If IsMissing(my_rng) Then Set my_rng = Selection
    
    For Each my_cell In my_rng
        my_cell = my_cell * -1
    Next my_cell

End Sub
