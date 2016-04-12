Option Explicit

Public Const SIZE = 8

Public b_chessboard(7, 7)               As Variant
Public l_solutions_found                As Long

Public attackedRows                     As Object ' as New Scripting.Dictionary => for early binding with Microsoft Scripting Runtime
Public attackedColumns                  As Object
Public attackedLeftDiagonals            As Object
Public attackedRightDiagonals           As Object

Sub Main()
    
    Set attackedRows = CreateObject("Scripting.Dictionary")
    Set attackedColumns = CreateObject("Scripting.Dictionary")
    Set attackedLeftDiagonals = CreateObject("Scripting.Dictionary")
    Set attackedRightDiagonals = CreateObject("Scripting.Dictionary")
    
    tbl_show.Cells.Delete
    l_solutions_found = 0
    Call PutQueens(0)
    tbl_show.Columns.ColumnWidth = 3
    
    Set attackedRows = Nothing
    Set attackedColumns = Nothing
    Set attackedLeftDiagonals = Nothing
    Set attackedRightDiagonals = Nothing
  
    
End Sub

Sub PutQueens(l_row As Long)
    
    Dim l_col        As Long
    
    If l_row = SIZE Then
        
        Call PrintSolution
        l_solutions_found = l_solutions_found + 1
        
    Else
        For l_col = 0 To SIZE - 1 Step 1
            If CanPlaceQueen(l_row, l_col) Then
                
                Call MarkAllAttackedPositions(l_row, l_col)
                Call PutQueens(l_row + 1)
                Call UnmarkAllattackedPositions(l_row, l_col)
            
            End If
        Next l_col
    End If
End Sub

Public Function CanPlaceQueen(l_row As Long, l_col As Long) As Boolean
    
    Dim b_result As Boolean
    
    b_result = dictionary_contains(attackedRows, l_row) Or _
                dictionary_contains(attackedColumns, l_col) Or _
                dictionary_contains(attackedLeftDiagonals, l_col - l_row) Or _
                dictionary_contains(attackedRightDiagonals, l_col + l_row)
    
    CanPlaceQueen = Not b_result
    
End Function

Public Sub PrintSolution()
    
    Dim l_row           As Long
    Dim l_col           As Long
    
    Dim l_row_fixer     As Long
    Dim l_col_fixer     As Long
    
    Dim s_result        As String
    
    l_row_fixer = (l_solutions_found \ 9) * 9 + 1
    l_col_fixer = (l_solutions_found Mod 9) * 9 + 1
 
    For l_row = 0 To SIZE - 1 Step 1
        For l_col = 0 To SIZE - 1 Step 1
            
            If b_chessboard(l_row, l_col) Then
                s_result = s_result & "*"
                tbl_show.Cells(l_row + l_row_fixer, l_col + l_col_fixer).Interior.Color = vbRed
            Else
                s_result = s_result & "-"
                tbl_show.Cells(l_row + l_row_fixer, l_col + l_col_fixer).Interior.Color = vbBlue
            End If
        Next l_col
        s_result = s_result & vbCrLf
    Next l_row
    
    Debug.Print l_solutions_found & vbCrLf & s_result
    
End Sub

Public Sub MarkAllAttackedPositions(l_row As Long, l_col As Long)
    
    attackedRows(l_row) = False
    attackedColumns(l_col) = False
    attackedLeftDiagonals(l_col - l_row) = False
    attackedRightDiagonals(l_col + l_row) = False
    
    b_chessboard(l_row, l_col) = True
    
End Sub

Public Sub UnmarkAllattackedPositions(l_row As Long, l_col As Long)
    
    attackedRows.Remove (l_row)
    attackedColumns.Remove (l_col)
    attackedLeftDiagonals.Remove (l_col - l_row)
    attackedRightDiagonals.Remove (l_col + l_row)
    
    b_chessboard(l_row, l_col) = False

End Sub

Public Function dictionary_contains(dict As Object, str_element As Variant) As Boolean
    
    Dim item        As Variant
    Dim b_result    As Boolean
    
    For Each item In dict
        If item = str_element Then b_result = True
    Next item
    
    dictionary_contains = b_result
    
End Function

Public Sub TestDictionary()
    
    attackedRows("a") = 1
    attackedRows("b") = 2
    attackedRows(15) = 3
    
    Debug.Print dictionary_contains(attackedRows, "b")
    Debug.Print dictionary_contains(attackedRows, "a")
    Debug.Print dictionary_contains(attackedRows, "d")
    Debug.Print dictionary_contains(attackedRows, "d")
    Debug.Print dictionary_contains(attackedRows, 15)
        
    Debug.Print "REMOVE"
    attackedRows.Remove ("a")
    Debug.Print dictionary_contains(attackedRows, "a")
    Debug.Print dictionary_contains(attackedRows, "a")
    
End Sub
