Public Function CollectionToArray(myCol As Collection) As Variant
    
    Dim result  As Variant
    Dim cnt     As Long
    
    ReDim result(myCol.Count - 1)

    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    CollectionToArray = result

End Function
