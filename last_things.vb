'locate last column 
'locate last row
'last things

Function last_col(Optional str_sheet As String, Optional row_to_check As Long = 1) As Long
    
    Dim shSheet  As Worksheet
    
        If str_sheet = vbNullString Then
            Set shSheet = ActiveSheet
        Else
            Set shSheet = Worksheets(str_sheet)
        End If
    
    last_col = shSheet.Cells(row_to_check, shSheet.Columns.Count).End(xlToLeft).Column

End Function


Function last_row(Optional str_sheet As String, Optional column_to_check As Long = 1) As Long
    
    Dim shSheet  As Worksheet
    
        If str_sheet = vbNullString Then
            Set shSheet = ActiveSheet
        Else
            Set shSheet = Worksheets(str_sheet)
        End If
    
    last_row = shSheet.Cells(shSheet.Rows.Count, column_to_check).End(xlUp).Row

End Function         
            
Public Function LastUsedColumn() As Long
    
    Dim rLastCell As Range
    
    Set rLastCell = ActiveSheet.Cells.Find(What:="*", _
                                    After:=ActiveSheet.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    
    LastUsedColumn = rLastCell.Column

End Function

Public Function LastUsedRow() As Long

    Dim rLastCell As Range

    Set rLastCell = ActiveSheet.Cells.Find(What:="*", _
                                    After:=ActiveSheet.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)

    LastUsedRow = rLastCell.Row

End Function

Public Function fnLngLocateValueRow(ByVal strTarget, _
                                    ByRef wksTarget As Worksheet, _
                                    Optional lngCol As Long = 1, _
                                    Optional lngMoreValuesFound As Long = 1, _
                                    Optional blnLookForPart = False, _
                                    Optional blnLookUpToBottom = True) As Long

    Dim lngValuesFound      As Long
    Dim rngLocal            As Range
    Dim rngMyCell           As Range
    
    fnLngLocateValueRow = -999
    lngValuesFound = lngMoreValuesFound
    Set rngLocal = wksTarget.Range(wksTarget.Cells(1, lngCol), wksTarget.Cells(Rows.Count, lngCol))

    For Each rngMyCell In rngLocal
        If blnLookForPart Then
            If strTarget = Left(rngMyCell, Len(strTarget)) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueRow = rngMyCell.row
                    If blnLookUpToBottom Then Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        Else
            If strTarget = Trim(rngMyCell) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueRow = rngMyCell.row
                    If blnLookUpToBottom Then Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        End If
    Next rngMyCell

End Function

Public Function fnLngLocateValueCol(ByVal strTarget, _
                                    ByRef wksTarget As Worksheet, _
                                    Optional lngRow As Long = 1, _
                                    Optional lngMoreValuesFound As Long = 1, _
                                    Optional blnLookForPart = False) As Long
                                    
    Dim lngValuesFound          As Long
    Dim rngLocal                As Range
    Dim rngMyCell               As Range
    
    lngValuesFound = lngMoreValuesFound
    Set rngLocal = wksTarget.Range(wksTarget.Cells(lngRow, 1), wksTarget.Cells(lngRow, Columns.Count))
    
    For Each rngMyCell In rngLocal
        If blnLookForPart Then
            If strTarget = Left(rngMyCell, Len(strTarget)) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueCol = rngMyCell.row
                    Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        Else
            If strTarget = Trim(rngMyCell) Then
                If lngValuesFound = 1 Then
                    fnLngLocateValueCol = rngMyCell.Column
                    Exit Function
                Else
                    Call Decrement(lngValuesFound)
                End If
            End If
        End If
    Next rngMyCell
    
    fnLngLocateValueCol = -999

End Function
