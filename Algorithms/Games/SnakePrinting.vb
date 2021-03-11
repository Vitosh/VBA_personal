Option Explicit

Public Function SnakeMyNumbers(n As Long) As String

    Dim lngCol As Long
    Dim lngRow As Long
    Dim str As String
    
    For lngCol = 0 To n - 1
    
        str = ""
        
        For lngRow = 0 To n - 1
            If lngRow Mod 2 = 0 Then
                str = str & vbTab & n * lngRow + lngCol + 1
            Else
                str = str & vbTab & n * (lngRow + 1) - lngCol
            End If
        Next lngRow
        
        SnakeMyNumbers = SnakeMyNumbers & str & vbCrLf
    Next lngCol

End Function
