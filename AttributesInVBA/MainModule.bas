Attribute VB_Name = "MainModule"
Option Explicit

Public Sub Main()
    
    'Because of
    '   Attribute VB_PredeclaredId = True
    'we can refer to CarGlobal without initialization:

    Debug.Print CarGlobal.Price
    Debug.Print CarGlobal.Model
    Debug.Print CarGlobal.ChangePrice(100)
    Debug.Print CarGlobal.Price
    
    'Because of
    '   Attribute Value.VB_Description = ""
    '   Attribute Value.VB_UserMemId = 0
    'the car has a a default property Price and it has description in the VBEditor
    
    Dim car As New CarWithDefaultProperty
    Debug.Print car
    
    'Because of
    '    Attribute Value.VB_UserMemId = 0
    '    Attribute Value.VB_Description = "Increases the price with 10. It is the default."

    Dim truck As New TruckWithDefaultProcedure
    Debug.Print truck.Price
    truck
    truck
    Debug.Print truck.Price
    
End Sub
