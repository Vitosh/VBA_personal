Sub FormatHalfOfTheSelectedCell()

    Dim myRange As Range
    Dim color As Long: color = RGB(0, 0, 0)
    Dim myShape As Shape
    
    With Worksheets("Sheet1") 'With ActiveSheet
    
        Set myRange = .Range("E10") 'Selection
        Dim left As Long: left = myRange.left
        Dim top As Long: top = myRange.top
        Dim width As Long: width = myRange.width
        Dim heigth As Long: heigth = myRange.Height

        'Top line:
        Set myShape = .Shapes.AddConnector(msoConnectorStraight, left, top, left + width / 2, top)
        myShape.Line.ForeColor.RGB = color
        
        'Left line:
        Set myShape = .Shapes.AddConnector(msoConnectorStraight, left, top, left, top + myRange.Height)
        myShape.Line.ForeColor.RGB = color
        
        'Right line:
        Set myShape = .Shapes.AddConnector(msoConnectorStraight, left + width / 2, top, left + width / 2, top + myRange.Height)
        myShape.Line.ForeColor.RGB = color
        
        Set myRange = myRange.Offset(1)
        left = myRange.left
        top = myRange.top
        width = myRange.width
        heigth = myRange.Height
                
        'Bottom line:
        Set myShape = .Shapes.AddConnector(msoConnectorStraight, left, top, left + width / 2, top)
        myShape.Line.ForeColor.RGB = RGB(200, 0, 0)

    End With

End Sub
