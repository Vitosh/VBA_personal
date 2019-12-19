Attribute VB_Name = "ExcelLastThings"
Option Explicit
Option Private Module

Public Function LastColumn(wsName As String, Optional rowToCheck As Long = 1) As Long

    Dim ws  As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    LastColumn = ws.Cells(rowToCheck, ws.Columns.Count).End(xlToLeft).Column
    
End Function

Public Function LastRow(wsName As String, Optional columnToCheck As Long = 1) As Long

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    LastRow = ws.Cells(ws.Rows.Count, columnToCheck).End(xlUp).Row

End Function
            
Public Function LastUsedColumn(wsName As String) As Long
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    Dim lastCell As Range
    
    Set lastCell = ActiveSheet.Cells.Find(What:="*", _
                                    After:=ActiveSheet.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    
    LastUsedColumn = lastCell.Column

End Function

Public Function LastUsedRow(wsName As String) As Long
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)
    Dim lastCell As Range

    Set lastCell = ActiveSheet.Cells.Find(What:="*", _
                                    After:=ActiveSheet.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

    LastUsedRow = lastCell.Row

End Function

Public Function LocateValueRow(ByVal textTarget As String, _
                ByRef wksTarget As Worksheet, _
                Optional col As Long = 1, _
                Optional moreValuesFound As Long = 1, _
                Optional lookForPart = False, _
                Optional lookUpToBottom = True) As Long

    Dim valuesFound      As Long
    Dim localRange            As Range
    Dim myCell           As Range

    LocateValueRow = -999
    valuesFound = moreValuesFound
    Set localRange = wksTarget.Range(wksTarget.Cells(1, col), wksTarget.Cells(Rows.Count, col))

    For Each myCell In localRange
        If lookForPart Then
            If textTarget = Left(myCell, Len(textTarget)) Then
                If valuesFound = 1 Then
                    LocateValueRow = myCell.Row
                    If lookUpToBottom Then Exit Function
                Else
                    Decrement valuesFound
                End If
            End If
        Else
            If textTarget = Trim(myCell) Then
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
    
    LocateValueCol = -999
    valuesFound = moreValuesFound
    Set localRange = wksTarget.Range(wksTarget.Cells(rowNeeded, 1), wksTarget.Cells(rowNeeded, Columns.Count))

    For Each myCell In localRange
        If lookForPart Then
            If textTarget = Left(myCell, Len(textTarget)) Then
                If valuesFound = 1 Then
                    LocateValueCol = myCell.Column
                    If lookUpToBottom Then Exit Function
                Else
                    Decrement valuesFound
                End If
            End If
        Else
            If textTarget = Trim(myCell) Then
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
                               
Public Sub Increment(ByRef valueToIncrement As Variant, Optional incrementWith As Double = 1)
    valueToIncrement = valueToIncrement + incrementWith
End Sub

Public Sub Decrement(ByRef valueToDecrement As Variant, Optional decrementWith As Double = 1)
    valueToDecrement = valueToDecrement - decrementWith
End Sub
