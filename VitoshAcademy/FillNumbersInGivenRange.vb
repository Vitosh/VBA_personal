Option Explicit

Private currentMove As Direction
Private size As Long

Public Enum Direction
    Right
    Down
    Left
    Up
End Enum

Sub Main()
    
    Cells.Clear
    size = 2
    SetMatrixStars
    MakeMatrix
    Cells.Columns.AutoFit

End Sub

Sub SetMatrixStars()
      
    
    Dim i As Long
    For i = 1 To size
        Cells(size + 1, i) = "*"
        Cells(i, size + 1) = "*"
    Next i
    
    Cells(size + 1, size + 1) = "*"
    
End Sub

Sub MakeMatrix()
    
    Dim currentCell As Range: Set currentCell = Cells(1, 1)
        
    currentMove = Right
    Dim i As Long

    Do While True
        i = i + 1
        currentCell = i
        If IsLast(currentCell) Then Exit Do
        Set currentCell = nextCell(currentCell)
    Loop
    
End Sub

Function IsLast(currentCell As Range) As Boolean
    
    If size = 1 Then
        IsLast = True
        Exit Function
    End If
    
    If currentCell.Row = 1 Or currentCell.Column = 1 Then
        If size = 2 And currentCell = 4 Then
            IsLast = True
        Else
            IsLast = False
        End If
        Exit Function
    End If
    
    IsLast = Not IsEmpty(currentCell.Offset(1, 0)) _
            And Not IsEmpty(currentCell.Offset(-1, 0)) _
            And Not IsEmpty(currentCell.Offset(0, -1)) _
            And Not IsEmpty(currentCell.Offset(0, 1))
    
End Function


Public Function nextCell(currentCell As Range) As Range
    
    Select Case currentMove
    
        Case Direction.Right
            If IsEmpty(currentCell.Offset(, 1)) Then
                Set nextCell = currentCell.Offset(, 1)
            Else
                Set nextCell = currentCell.Offset(1)
                currentMove = Direction.Down
            End If
            
        Case Direction.Down
            If IsEmpty(currentCell.Offset(1)) Then
                Set nextCell = currentCell.Offset(1)
            Else
                Set nextCell = currentCell.Offset(, -1)
                currentMove = Direction.Left
            End If
            
        Case Direction.Left
            If currentCell.Column = 1 Then
                Set nextCell = currentCell.Offset(-1)
                currentMove = Direction.Up
            Else
                If IsEmpty(currentCell.Offset(, -1)) Then
                    Set nextCell = currentCell.Offset(, -1)
                Else
                    Set nextCell = currentCell.Offset(-1)
                    currentMove = Direction.Up
                End If
            End If
            
        Case Direction.Up
            If IsEmpty(currentCell.Offset(-1)) Then
                Set nextCell = currentCell.Offset(-1)
            Else
                Set nextCell = currentCell.Offset(0, 1)
                currentMove = Direction.Right
            End If
    End Select
    
End Function
