Option Explicit

Sub ShapeNames()
    Dim sh_shape As shape
    
    For Each sh_shape In ActiveSheet.Shapes
        Debug.Print sh_shape.Name
    Next sh_shape
    
End Sub

Public Sub GetSomething(str_something As String)
    
    ActiveSheet.Shapes(str_something).Select

End Sub

Option Explicit

'Makes shape visible and invisble.
Sub translatorField_Klicken()

    Dim blnEnglish      As Boolean
    Dim rngRange        As Range
    Dim myShape         As shape

    Set myShape = tblInput.Shapes("translatorField")
    Set rngRange = tblSettings.Cells(2, 2)

    blnEnglish = Not CBool(rngRange)
    tblSettings.Cells(2, 2) = blnEnglish

    If blnEnglish Then

        tblInput.[h1].value = tblSettings.[i1].value

        With myShape.Fill
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 1
        End With

        With myShape.TextFrame2.TextRange.Characters(1, 66).Font.Fill
            .ForeColor.RGB = RGB(255, 255, 255)
            .Transparency = 1
        End With

    Else

        tblInput.[h1].value = tblSettings.[c1].value

        With myShape.Fill
            .ForeColor.RGB = RGB(255, 255, 255)
            .Transparency = 0
        End With

        With myShape.TextFrame2.TextRange.Characters(1, 66).Font.Fill
            .ForeColor.RGB = RGB(0, 0, 0)
            .Transparency = 0
        End With

    End If

End Sub
