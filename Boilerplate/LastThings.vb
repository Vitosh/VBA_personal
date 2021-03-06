Option Explicit
Option Private Module
    
'locate last column 
'locate last row
'last things count substrings, count strings, count stuff

Public Function LastColumn(ws As Worksheet, Optional rowToCheck As Long = 1) As Long

    LastColumn = ws.Cells(rowToCheck, ws.Columns.count).End(xlToLeft).Column
    
End Function

Public Function LastRow(ws As Worksheet, Optional columnToCheck As Long = 1) As Long
    
    LastRow = ws.Cells(ws.Rows.count, columnToCheck).End(xlUp).Row

End Function
            
            
Public Function LastUsedColumn(wks As Worksheet) As Long
    
    Dim lastCell As Range
    
    With wks
        Set lastCell = .Cells.Find(What:="*", _
                    After:=.Cells(1, 1), _
                    LookIn:=xlFormulas, _
                    LookAt:=xlPart, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False)
    End With    
    LastUsedColumn = lastCell.Column

End Function


Public Function LocateValueRow(ByVal textTarget As String, _
                ByRef wksTarget As Worksheet, _
                Optional col As Long = 1, _
                Optional moreValuesFound As Long = 1, _
                Optional lookForPart = False, _
                Optional lookUpToBottom = True) As Long

    Dim valuesFound         As Long
    Dim localRange          As Range
    Dim myCell              As Range
    Dim lastRowOnColumn1    As Long
    
    LocateValueRow = GENERAL_NUMBERS.NF
    
    valuesFound = moreValuesFound
    lastRowOnColumn1 = LastRow(wksTarget, col)
    Set localRange = wksTarget.Range(wksTarget.Cells(1, col), wksTarget.Cells(lastRowOnColumn1, col))

    For Each myCell In localRange
        If lookForPart Then
            If UCase(textTarget) = UCase(Left(myCell, Len(textTarget))) Then
                If valuesFound = 1 Then
                    LocateValueRow = myCell.Row
                    If lookUpToBottom Then Exit Function
                Else
                    Decrement valuesFound
                End If
            End If
        Else
            If UCase(textTarget) = UCase(Trim(myCell)) Then
                If valuesFound = 1 Then
                    LocateValueRow = myCell.Row
                    If lookUpToBottom Then Exit Function
                Else
                    Decrement valuesFound
                End If
            End If
        End If
    Next myCell

End Function

Public Function LocateValueCol(ByVal textTarget As String, _
                ByRef wksTarget As Worksheet, _
                Optional rowNeeded As Long = 1, _
                Optional moreValuesFound As Long = 1, _
                Optional lookForPart = False, _
                Optional lookUpToBottom = True) As Long

    Dim valuesFound As Long
    Dim localRange  As Range
    Dim myCell  As Range
    
    LocateValueCol = GENERAL_NUMBERS.NF
    valuesFound = moreValuesFound
    Set localRange = wksTarget.Range(wksTarget.Cells(rowNeeded, 1), wksTarget.Cells(rowNeeded, Columns.count))

    For Each myCell In localRange
        If lookForPart Then
            If UCase(textTarget) = UCase(Left(myCell, Len(textTarget))) Then
                If valuesFound = 1 Then
                    LocateValueCol = myCell.Column
                    If lookUpToBottom Then Exit Function
                Else
                    Decrement valuesFound
                End If
            End If
        Else
            If UCase(textTarget) = UCase(Trim(myCell)) Then
                If valuesFound = 1 Then
                    LocateValueCol = myCell.Column
                    If lookUpToBottom Then Exit Function
                Else
                    Decrement valuesFound
                End If
            End If
        End If
    Next myCell

End Function
                
                
Public Function GetColumnSequence(tbl As Worksheet, tableName As String, columnName As String) As Long
        
    Dim myCell As Range
    Dim result As Long
    result = 1
    
    For Each myCell In ThisWorkbook.Worksheets(tbl.Name).Range(tableName & "[#Headers]").Cells
        If UCase(Trim(myCell)) = UCase(Trim(columnName)) Then
            GetColumnSequence = result
            Exit Function
        Else
            result = result + 1
        End If
    Next
    
    GetColumnSequence = -1
    
End Function
            
                
Private Sub Increment(ByRef valueToIncrement As Variant, Optional incrementWith As Double = 1)
    
    valueToIncrement = valueToIncrement + incrementWith

End Sub

Private Sub Decrement(ByRef valueToDecrement As Variant, Optional decrementWith As Double = 1)

    valueToDecrement = valueToDecrement - decrementWith

End Sub
                
Public Function CountSubstringsInRow(wks As Worksheet, substring As String, Optional myRow As Long = 1)
        
    Dim myLastCol As Long
    myLastCol = LastColumn(wks, myRow)
    
    Dim result As Long
    Dim myCell As Range
    
    With wks
        For Each myCell In .Range(.Cells(myRow, 1), .Cells(myRow, myLastCol))
            If InStr(1, myCell.Text, substring, vbTextCompare) Then
                result = result + 1
            End If
        Next
    End With
    
    CountSubstringsInRow = result
    
End Function

                    
'LastRow Last Row Formula
=IFERROR(LOOKUP(2,1/(NOT(ISBLANK(A:A))),ROW(A:A)),0)

'LastColumn Last Column Formula
=IFERROR(LOOKUP(2,1/(NOT(ISBLANK(1:1))),COLUMN(1:1)),0)
                                    
'Last Row Value of Column A
=LOOKUP(2,1/(NOT(ISBLANK(A:A))),A:A)
                                    
'Last Column Value of the first row
=LOOKUP(2,1/(NOT(ISBLANK(1:1))),1:1)

