Public Sub AddCommentToSelection(my_comment As Range)
  
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

'Make Comments even better:

Public Sub AddComments(r_cell As Range)
    
    r_cell.ClearComments
    r_cell.AddComment.Visible = False
    r_cell.Comment.Text (generate_info_for_comment(r_cell))
    
    With r_cell.Comment.Shape
        
        .AutoShapeType = msoShapeRoundedRectangle
        
        .ScaleHeight 3.5, msoFalse, msoScaleFromTopLeft
        .ScaleWidth 4, msoFalse, msoScaleFromTopLeft
        
        .TextFrame.Characters.Font.Name = "Tahoma"
        .TextFrame.Characters.Font.Size = 14
        .TextFrame.Characters.Font.ColorIndex = 1

        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.BackColor.RGB = RGB(255, 255, 255)
        
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 204, 153)
        .Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.25
        '.Fill.OneColorGradient msoGradientDiagonalUp, 2, 0.9
        '.Fill.TwoColorGradient msoGradientHorizontal, 2

        .Line.DashStyle = msoLineLongDash        
        .Shadow.Visible = msoFalse
        
        .Placement = xlMoveAndSize
        
    End With
    
End Sub

Public Function generate_info_for_comment(my_cell As Range) As String
    
    Dim str_text As String
    
    str_text = "Auto " & Left(Date, 5) & " " & Left(Environ("username"), 4) & vbCrLf & vbCrLf
    str_text = str_text & "Werte:" & " " & my_cell.value & vbCrLf & vbCrLf
    str_text = str_text & "war:" & " " & my_cell.Formula
        
    generate_info_for_comment = str_text
    
End Function

Public Sub FixComments()

    Dim xComment As Comment
    
    For Each xComment In Application.ActiveSheet.Comments
        xComment.Shape.TextFrame.AutoSize = True
    Next xComment
    
End Sub

