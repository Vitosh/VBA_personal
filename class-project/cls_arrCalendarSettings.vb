Option Explicit

Private p_top_row                           As Long
Private p_bottom_row                        As Long
Private p_left_col                          As Long
Private p_right_col                         As Long

Private p_sonstiges_pro_BA                  As Double
Private p_verhaltnis_baukosten_planer       As Double

Private p_vertriebsstart                    As Date
Private p_vertriebsstart_col_num            As Long

Public Property Let Vertriebsstart_Col(l_vertriebsstart_col As Long)
    p_vertriebsstart_col_num = l_vertriebsstart_col
End Property

Public Property Get Vertriebsstart_Col() As Long
    Vertriebsstart_Col = p_vertriebsstart_col_num
End Property

Public Property Let Vertriebsstart(date_vertriebsstart As Date)
    p_vertriebsstart = date_vertriebsstart
End Property

Public Property Get Vertriebsstart() As Date
    Vertriebsstart = p_vertriebsstart
End Property

Public Property Get LengthLeftToRight() As Long
    LengthLeftToRight = RightCol - LeftCol
End Property

Public Property Get LengthTopToBottom() As Long
    LengthTopToBottom = BottomRow - TopRow
End Property

Public Property Let VerhaltnisBaukostenToPlanerkosten(dbl_verhaltnis As Double)
    p_verhaltnis_baukosten_planer = dbl_verhaltnis
End Property

Public Property Get VerhaltnisBaukostenToPlanerkosten() As Double
    VerhaltnisBaukostenToPlanerkosten = p_verhaltnis_baukosten_planer
End Property

Public Property Let SonstigesProBA(dbl_sonstiges_money As Double)
    p_sonstiges_pro_BA = dbl_sonstiges_money
End Property

Public Property Get SonstigesProBA() As Double
    SonstigesProBA = p_sonstiges_pro_BA
End Property

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
