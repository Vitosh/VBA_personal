Public Sub add_comment_to_selection(my_comment As Range)
  
  'b is used only if the cells are merged by 2.
  
    Dim b As Boolean
    b = True
    For Each current_cell In selection
        If b Then
            current_cell.ClearComments
            current_cell.AddComment my_comment.Text
            current_cell.Comment.Visible = False
            current_cell.Comment.Shape.ScaleWidth 4, msoFalse, msoScaleFromTopLeft
            current_cell.Comment.Shape.ScaleHeight 2.26, msoFalse, msoScaleFromTopLeft
        End If
        'b = Not b
    Next current_cell
End Sub

Public Sub delete_comment_in_selection()
    For Each current_cell In selection
        current_cell.ClearComments
    Next current_cell
End Sub
