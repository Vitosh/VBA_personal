Sub HideShowComments(Optional b_show_comments As Boolean = False)
    On Error Resume Next
    For Each current_cell In Range("A1:AO1000")
        current_cell.Comment.Visible = b_show_comments
    Next current_cell
    On Error GoTo 0
End Sub

