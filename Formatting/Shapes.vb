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

'---------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Sub TestMe()

    Dim shp             As Shape
    Dim arrOfShapes()   As Variant

    With ActiveSheet
        For Each shp In .Shapes
            If InStrB(shp.Name, "Rec") > 0 Then
                arrOfShapes = incrementArray(arrOfShapes, shp.Name)
            End If
        Next
        If IsArrayAllocated(arrOfShapes) Then
            Debug.Print .Shapes.Range(arrOfShapes(0)).Name
            .Shapes.Range(arrOfShapes).Delete
        End If
    End With
End Sub


Public Function incrementArray(arrOfShapes As Variant, nameOfShape As String) As Variant

    Dim cnt         As Long
    Dim arrNew      As Variant

    If IsArrayAllocated(arrOfShapes) Then
        ReDim arrNew(UBound(arrOfShapes) + 1)            
        For cnt = LBound(arrOfShapes) To UBound(arrOfShapes)
            arrNew(cnt) = CStr(arrOfShapes(cnt))
        Next cnt
        arrNew(UBound(arrOfShapes) + 1) = CStr(nameOfShape)
    Else
        arrNew = Array(nameOfShape)
    End If

    incrementArray = arrNew

End Function

Function IsArrayAllocated(Arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(Arr) And _
                       Not IsError(LBound(Arr, 1)) And _
                       LBound(Arr, 1) <= UBound(Arr, 1)

End Function
Credits to this guy for the finding that the arrOfShapes should be declared with parenthesis (I have spent about 30 minutes researching why I could not pass it correctly) and to CPearson for the IsArrayAllocated().

