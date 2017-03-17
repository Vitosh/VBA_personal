Attribute VB_Name = "mod_main"
Option Explicit

Public obj_project As cls_project

Public Sub SetObjectBA()

    Dim l_counter   As Long
    
    Set obj_project = New cls_project
    
    For l_counter = 0 To 2
        obj_project.AddBA Cls_BA_Builder(l_counter, l_counter + 5, Now())
    Next l_counter

End Sub

Public Function Cls_BA_Builder(f_count_ba As Long, _
                                f_row As Long, _
                                f_vertriebsstart As Date) As cls_ba


    Dim obj                 As cls_ba
    
    Set obj = New cls_ba
    
    obj.CounterBA = f_count_ba
    obj.Row = f_row
    obj.Vertriebsstart = f_vertriebsstart
    
    Set Cls_BA_Builder = obj

End Function


