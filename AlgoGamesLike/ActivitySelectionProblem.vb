Option Explicit

Public Sub TestMe()

    Dim objA            As clsActivity
    Dim colObjs         As New Collection
    Dim rngCell         As Range
    Dim strResult       As String
    Dim i               As Long
    Dim lngNextStart    As Long: lngNextStart = 0
    
    For Each rngCell In Range(Cells(1, 1), Cells(1, 11))
        Set objA = Nothing
        Set objA = New clsActivity
        objA.StartTime = rngCell
        objA.EndTime = rngCell.Offset(1, 0)
        objA.Name = rngCell.Offset(2, 0)
        colObjs.Add objA
    Next rngCell
    
    Set colObjs = SortedCollection(colObjs)
    
    For i = 1 To colObjs.Count
        If colObjs.Item(i).StartTime > lngNextStart Then
            strResult = strResult & colObjs.Item(i).Name & vbTab & _
                                    colObjs.Item(i).StartTime & vbTab & _
                                    colObjs.Item(i).EndTime & vbCrLf
                                    
            lngNextStart = colObjs.Item(i).EndTime
        End If
    Next i
    
    Debug.Print strResult
    
End Sub

Public Function SortedCollection(myColl As Collection, Optional blnSortABC As Boolean = True) As Collection

    Dim i           As Long
    Dim j           As Long
    
    For i = myColl.Count To 2 Step -1
        For j = 1 To i - 1
            If blnSortABC Then
                If myColl(j).EndTime > myColl(j + 1).EndTime Then
                    myColl.Add myColl(j), after:=j + 1
                    myColl.Remove j
                End If
            Else
                If myColl(j).EndTime < myColl(j + 1).EndTime Then
                    myColl.Add myColl(j), after:=j + 1
                    myColl.Remove j
                End If
            End If
        Next j
    Next i
    
    Set SortedCollection = myColl
    

End Function

