Option Explicit

Private p_top_row                           As Long
Private p_bottom_row                        As Long
Private p_left_col                          As Long
Private p_right_col                         As Long

Public Property Let TopRow(l_top_row As Long)
    p_top_row = l_top_row
End Property

Public Property Get TopRow() As Long
    TopRow = p_top_row
End Property

Public Property Let BottomRow(l_bottom_row As Long)
    p_bottom_row = l_bottom_row
End Property

Public Property Get BottomRow() As Long
    BottomRow = p_bottom_row
End Property

Public Property Let LeftCol(l_left_col As Long)
    p_left_col = l_left_col
End Property

Public Property Get LeftCol() As Long
    LeftCol = p_left_col
End Property

Public Property Let RightCol(l_right_col As Long)
    p_right_col = l_right_col
End Property

Public Property Get RightCol() As Long
    RightCol = p_right_col
End Property
