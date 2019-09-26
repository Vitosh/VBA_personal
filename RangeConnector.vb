Sub FormatHalfOfTheSelectedCell()

    Dim myRange As Range
    Set myRange = Selection
    
    Dim l As Long
    Dim t As Long
    Dim w As Long
    Dim h As Long
    
    l = myRange.Left
    t = myRange.Top
    w = myRange.Width
    h = myRange.Height
    
    ActiveSheet.Shapes.AddConnector msoConnectorStraight, l, t, l + (w) / 2, t
    ActiveSheet.Shapes.AddConnector msoConnectorStraight, l, t, l, t + myRange.Height
    
    Set myRange = myRange.Offset(1)
    l = myRange.Left
    t = myRange.Top
    w = myRange.Width
    h = myRange.Height
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, l, t, l + (w) / 2, t).Select
    

    myRange.Select

End Sub
