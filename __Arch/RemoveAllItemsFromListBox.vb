Private Sub RemoveAllItemsFromListBox(lb_object As Object)
    
    Dim l_counter   As Long
    
    For l_counter = 1 To lb_object.ListCount
        lb_object.RemoveItem 0
    Next l_counter

End Sub
