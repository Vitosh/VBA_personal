Option Explicit

'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646299(v=vs.85).aspx
'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646293(v=vs.85).aspx

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long

Private Const SIZE_WIDTH            As Long = 7
Private Const SIZE_HEIGTH           As Long = 5
Private Const COL_WIDTH             As Double = 2.3
Private Const BORDER_COL            As Long = 190

Private wks                         As Worksheet
Private pointX                      As Long
Private pointY                      As Long
Private leadPoint                   As Range
Private pointField                  As Range

Private movingDirection             As Direction
Public Enum Direction

    GoUp = 1
    GoRight = 2
    GoDown = 3
    GoLeft = 4

End Enum

Private Sub Main()
    
    FixThePitch
    InitializePoint
    PrintInformation
    MoveAround
    
End Sub

Public Sub PrintInformation()
    
    Debug.Print "Press Home to exit."
    
End Sub

Private Sub ShowNewFood()
    
    Dim positionRow         As Long
    Dim positionCol         As Long
    
    positionRow = 1
    positionCol = 1
    
End Sub

Private Function MakeRandom(down As Long, up As Long) As Long

    MakeRandom = CLng((up - down) * Rnd + down)

End Function

Public Sub ChangePoints(pointToChange As Long)

    pointField.value = pointField + pointToChange

End Sub

Public Sub GoMove(moveDir As Direction)
    
    Debug.Print moveDir
    
End Sub

Public Sub ColorSnake()
    
    With wks
        .Range(.Cells(1, 1), .Cells(SIZE_HEIGTH, SIZE_WIDTH)).Clear
    End With
    leadPoint.Interior.COLOR = vbWhite

End Sub

Private Sub MoveFurther()
    
    Select Case movingDirection
    
        Case GoUp:
            If leadPoint.row = 1 Then
                Set leadPoint = Cells(SIZE_HEIGTH, leadPoint.Column)
            Else
                Set leadPoint = Cells(leadPoint.row - 1, leadPoint.Column)
            End If
            
        Case GoRight:
            If leadPoint.Column = SIZE_WIDTH Then
                Set leadPoint = Cells(leadPoint.row, 1)
            Else
                Set leadPoint = Cells(leadPoint.row, leadPoint.Column + 1)
            End If
        
        Case GoDown:
            If leadPoint.row = SIZE_HEIGTH Then
                Set leadPoint = Cells(1, leadPoint.Column)
            Else
                Set leadPoint = Cells(leadPoint.row + 1, leadPoint.Column)
            End If
        
        Case GoLeft:
            If leadPoint.Column = 1 Then
                Set leadPoint = Cells(leadPoint.row, SIZE_WIDTH)
            Else
                Set leadPoint = Cells(leadPoint.row, leadPoint.Column - 1)
            End If
    End Select
    
End Sub

Private Sub ReadKey()

    Debug.Assert Not IsEmpty(GetAsyncKeyState(vbKeyUp))
    
    Select Case True
        Case GetAsyncKeyState(vbKeyHome)
            Debug.Print "Exiting..."
            End
            
        Case GetAsyncKeyState(vbKeyUp):
            movingDirection = GoUp
            
        Case GetAsyncKeyState(vbKeyRight):
            movingDirection = GoRight
            
        Case GetAsyncKeyState(vbKeyDown):
            movingDirection = GoDown
                    
        Case GetAsyncKeyState(vbKeyLeft):
            movingDirection = GoLeft
    End Select
    
End Sub

Private Sub MoveAround()

    movingDirection = Direction.GoRight
    
    Do While True
        DoEvents
        ReadKey
        ColorSnake
        MoveFurther
        Sleep (200)
    Loop

End Sub

Private Sub InitializePoint()

    Set leadPoint = wks.Cells(2, 3)

End Sub

Private Sub FixThePitch()

    Set wks = tbl_Internal1

    wks.visible = xlSheetVisible
    wks.Activate
    
    With wks
        .Cells.Delete
        .Cells(1, 1).Select
        .Range(.Cells(1), .Cells(1 + SIZE_WIDTH)).ColumnWidth = COL_WIDTH
        .Range(.Cells(SIZE_HEIGTH + 1, 1), .Cells(SIZE_HEIGTH + 1, SIZE_WIDTH)).Borders.COLOR = RGB(BORDER_COL, BORDER_COL, BORDER_COL)
        .Range(.Cells(1, SIZE_WIDTH + 1), .Cells(SIZE_HEIGTH + 1, SIZE_WIDTH + 1)).Borders.COLOR = RGB(BORDER_COL, BORDER_COL, BORDER_COL)
    End With

    Set pointField = wks.Cells(8, 1)
    ChangePoints (0)
    
End Sub
