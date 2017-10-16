Option Explicit

'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646299(v=vs.85).aspx
'https://msdn.microsoft.com/en-us/library/windows/desktop/ms646293(v=vs.85).aspx

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SIZE_WIDTH    As Long = 7
Private Const SIZE_HEIGTH   As Long = 5
Private Const COL_WIDTH     As Double = 2.3

Private wks                 As Worksheet
Private pointX              As Long
Private pointY              As Long
Private leadPoint           As Range

Private movingDirection     As Direction

Public Enum Direction
    GoUp = 1
    GoRight = 2
    GoDown = 3
    GoLeft = 4
End Enum

Private Sub Main()
    
    FixThePitch
    InitializePoint
    MoveAround
    
End Sub

Public Sub GoMove(moveDir As Direction)

    Debug.Print moveDir

End Sub

Public Sub ColorSnake()

    wks.Cells.Clear
    leadPoint.Interior.Color = vbBlue

End Sub

Private Sub MoveFurther()
    
    Select Case movingDirection
    
        Case GoUp:
            If leadPoint.Row = 1 Then
                Set leadPoint = Cells(SIZE_HEIGTH, leadPoint.Column)
            Else
                Set leadPoint = Cells(leadPoint.Row - 1, leadPoint.Column)
            End If
            
        Case GoRight:
            If leadPoint.Column = SIZE_WIDTH Then
                Set leadPoint = Cells(leadPoint.Row, 1)
            Else
                Set leadPoint = Cells(leadPoint.Row, leadPoint.Column + 1)
            End If
        
        Case GoDown:
            If leadPoint.Row = SIZE_HEIGTH Then
                Set leadPoint = Cells(1, leadPoint.Column)
            Else
                Set leadPoint = Cells(leadPoint.Row + 1, leadPoint.Column)
            End If
        
        Case GoLeft:
            If leadPoint.Column = 1 Then
                Set leadPoint = Cells(leadPoint.Row, SIZE_WIDTH)
            Else
                Set leadPoint = Cells(leadPoint.Row, leadPoint.Column - 1)
            End If
    End Select
    
End Sub

Private Sub ReadKey()

    Select Case True
        
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
    Dim cnt As Long
    
    Do While True
        DoEvents
        ReadKey
        ColorSnake
        MoveFurther
        Sleep (500)
        cnt = cnt + 1
        Debug.Assert cnt < 20
    Loop

End Sub

Private Sub InitializePoint()

    Set leadPoint = wks.Cells(2, 3)

End Sub

Private Sub FixThePitch()

    Set wks = ActiveSheet
    
    wks.visible = xlSheetVisible
    wks.Activate
    
    With wks
        .Cells.Delete
        .Cells(1, 1).Select
        .Range(.Cells(1), .Cells(SIZE_WIDTH)).ColumnWidth = COL_WIDTH
    End With

End Sub

