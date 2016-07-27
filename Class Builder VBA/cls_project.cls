VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_ba_info()   As cls_ba

Public Sub AddBA(ByVal obj_ba As cls_ba)
        
    ReDim Preserve p_ba_info(UBound(p_ba_info) + 1)
    Set p_ba_info(UBound(p_ba_info)) = obj_ba

End Sub

Private Sub Class_Initialize()
    
    ReDim p_ba_info(0)
    
End Sub

