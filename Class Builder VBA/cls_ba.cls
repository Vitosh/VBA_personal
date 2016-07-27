VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_ba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_row                   As Long
Private p_vertriebsstart        As Date
Private p_counter_ba            As Long

Public Property Let Row(value As Long)
    p_row = value
End Property

Public Property Get Row() As Long
    Row = p_row
End Property

Public Property Let Vertriebsstart(value As Date)
    p_vertriebsstart = value
End Property

Public Property Get Vertriebsstart() As Date
    Vertriebsstart = p_vertriebsstart
End Property

Public Property Let CounterBA(value As Long)
    p_counter_ba = value
End Property

Public Property Get CounterBA() As Long
    CounterBA = p_counter_ba
End Property

