Public Sub AddCommentToSelection(myComment As String)
    
    Dim myCell As Range
    
    For Each myCell In Selection
        myCell.ClearComments
        myCell.AddComment myComment
        myCell.Comment.Visible = False
        myCell.Comment.Shape.ScaleWidth 4, msoFalse, msoScaleFromTopLeft
        myCell.Comment.Shape.ScaleHeight 2.26, msoFalse, msoScaleFromTopLeft
    Next myCell
    
End Sub

Public Sub DeleteCommentFromSelection()
    
    Dim myCell As Range
    
    For Each myCell In Selection
        myCell.ClearComments
    Next myCell
    
End Sub

Public Sub BeautifyComments(myCell As Range)
    
    myCell.ClearComments
    myCell.AddComment.Visible = False
    myCell.Comment.Text (GenerateInfoForComment(myCell))
    
    With myCell.Comment.Shape
        
        .AutoShapeType = msoShapeRoundedRectangle
        
        .ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft
        .ScaleWidth 2, msoFalse, msoScaleFromTopLeft
        
        .TextFrame.Characters.Font.Name = "Tahoma"
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.ColorIndex = 1

        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.BackColor.RGB = RGB(255, 255, 255)
        
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 204, 153)
        .Fill.OneColorGradient msoGradientDiagonalUp, 1, 0.25
        .Line.DashStyle = msoLineLongDash
        .Shadow.Visible = msoFalse
        .Placement = xlMoveAndSize
        
    End With
    
End Sub

Public Sub MakeAllCommentsVisible()

    Dim myComment As Comment

    For Each myComment In Application.ActiveSheet.Comments
        myComment.Visible = False
    Next myComment

End Sub