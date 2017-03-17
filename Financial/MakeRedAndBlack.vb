'---------------------------------------------------------------------------------------
' Module    : mod_main
' Author    : v.doynov
' Date      : 27.01.2016
' Purpose   : To make the tool work, we need four lines of values. Rows 1,2 and Rows 4,5
'             We need to put values only on row 4, positive and negative.
'             Then run the "main" procedure.
'---------------------------------------------------------------------------------------

Option Explicit

Public Const STARTING_FROM_COLUMN = 1
Public Const COLUMNS_NOT_TOUCHED = 0
Public current_cell                 As Range
'

Public Sub Main()
    
    Dim my_cell         As Range
    Dim l_col_len       As Long: l_col_len = last_column(row_to_check:=4)
    Dim l_counter       As Long
    Dim d_result        As Double
    Dim d_result_ini    As Double
    
    On Error GoTo main_Error
    
    Call OnStart
    
    tbl_output.Unprotect "toughpassword100"
    
    tbl_output.Rows(1).Clear
    tbl_output.Rows(2).Clear
    tbl_output.Rows(3).Clear
    tbl_output.Rows(5).Clear
    tbl_output.Rows(6).Clear
        
    'Copy
    Range(Cells(1, 1), Cells(1, l_col_len)).Value = Range(Cells(4, 1), Cells(4, l_col_len)).Value
    
    'Format
    Call MakeRedAndBlack(tbl_output.Cells(2, 1))
    Call MakeRedAndBlack(tbl_output.Cells(5, 1))
    
    Set my_cell = tbl_output.Cells(2, 1)
    my_cell.FormulaR1C1 = "=R[-1]C"
    my_cell.Offset(1, 0).Interior.Color = 5296274
    Call MakeRedAndBlack(my_cell)
    Call MakeRedAndBlack(my_cell.Offset(-1, 0))
    
    Set my_cell = tbl_output.Cells(5, 1)
    my_cell.FormulaR1C1 = "=R[-1]C"
    my_cell.Offset(1, 0).Interior.Color = 5296274
    Call MakeRedAndBlack(my_cell)
    Call MakeRedAndBlack(my_cell.Offset(-1, 0))
    
    For l_counter = 2 To l_col_len
    
        Set my_cell = tbl_output.Cells(2, l_counter)
        
        my_cell.Formula = "=R[-1]C+RC[-1]"
        my_cell.Offset(3, 0).Formula = "=R[-1]C+RC[-1]"
        
        my_cell.Offset(1, 0).Interior.Color = 5296274
        my_cell.Offset(4, 0).Interior.Color = 5296274
        
        Call MakeRedAndBlack(my_cell)
        Call MakeRedAndBlack(my_cell.Offset(-1, 0))
        Call MakeRedAndBlack(my_cell.Offset(2, 0))
        Call MakeRedAndBlack(my_cell.Offset(3, 0))
                        
    Next l_counter
    
    'Action
    Call RedAndBlackRecalculation_main2(l_col_len, 2)
    
    'Checks
    d_result = sum_range(tbl_output.Range(tbl_output.Cells(4, 1), tbl_output.Cells(4, l_col_len)))
    d_result_ini = sum_range(tbl_output.Range(tbl_output.Cells(1, 1), tbl_output.Cells(1, l_col_len)))
    
    If d_result > 0 Then
        [my_result] = d_result
        'MsgBox "Sie haben keinen Gewinn. Ihre finanziellen Verlust beträgt " & d_result & " Euro.", vbInformation, "RedAndBlack"
    Else
        [my_result] = ""
    End If
    
    'tbl_output.Protect "toughpassword100"
    
    If d_result <> d_result_ini Then
        MsgBox "Überprüfen Sie die Eingabe.", vbInformation, "RedAndBlack"
    End If
        
    tbl_output.Rows(2).EntireRow.Hidden = 1
    tbl_output.Rows(5).EntireRow.Hidden = 1
        
    Call OnEnd
    Set my_cell = Nothing

    On Error GoTo 0
    Exit Sub

main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main of Module mod_main"
    Call OnEnd
    
End Sub

Public Sub MakeRedAndBlack(ByRef my_range As Range)

    my_range.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
    my_range.Font.Name = "Calibri"
    my_range.Font.Size = 11
    
        
    'if we try to do it with parenthesis, then the zero values are not showing...
    'my_range.NumberFormat = "$#,##0.00_);[Red]($#,##0.00);"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RedAndBlackRecalculation
' Author    : v.doynov
' Date      : 07.08.2015
' Purpose   : Divides the row of "CashFlow vor Steuern" into red to the right and black
'             to the left. Change "calendar_cols" and "current_row" to make it work.
'             In order to call it use "call RedAndBlackRecalculation(27,84)".
'             84 is the middle line of the original 3 in PAKU.
'---------------------------------------------------------------------------------------
'
Public Sub RedAndBlackRecalculation_main2(ByVal calendar_cols As Long, ByVal current_row As Long)
    
    Dim counter                 As Long
    
    Dim final_col_in_loop       As Long
    
    Dim cell                    As Range
    Dim range_for_analysis      As Range
    
    Dim holdback                As Double
    Dim max_for_break_even      As Double
    
    Dim cell_with_break_even    As Range
    
    On Error GoTo RedAndBlackRecalculation_Error

    holdback = 0
    
    'When used outside PAKU remove "tbl_output.Range" for the set
    With tbl_output
        Set range_for_analysis = .Range(.Cells(current_row, STARTING_FROM_COLUMN), .Cells(current_row, calendar_cols + COLUMNS_NOT_TOUCHED))
    End With
    
    max_for_break_even = Application.WorksheetFunction.Max(range_for_analysis)
    
    For Each cell In range_for_analysis
        If cell.Value = max_for_break_even Then
            Set cell_with_break_even = cell
            Exit For
        End If
    Next cell
    
    final_col_in_loop = cell_with_break_even.Column + 1
    current_row = current_row - 1
    
    If cell_with_break_even.Column = 1 And cell_with_break_even <= 0 Then
        For counter = COLUMNS_NOT_TOUCHED + calendar_cols To cell_with_break_even.Column Step -1
        
            With tbl_output
                Set current_cell = .Cells(current_row, counter)
            End With
            
            If current_cell > 0 Then
                holdback = holdback + current_cell
                current_cell = 0
            Else
                current_cell = current_cell + holdback
                holdback = 0
            End If
            
            'we do it for a second time,
            'in order to make it equal to zero, if
            'it is not in the break even point
            
            If current_cell > 0 Then
                holdback = holdback + current_cell
                current_cell = 0
            End If
        Next counter
    Else
    
        For counter = COLUMNS_NOT_TOUCHED + calendar_cols To final_col_in_loop Step -1
            
            With tbl_output
                Set current_cell = .Cells(current_row, counter)
            End With
            
            If current_cell > 0 Then
                holdback = holdback + current_cell
                current_cell.Value = 0
            Else
                current_cell = current_cell + holdback
                holdback = 0
            End If
            
            'we do it for a second time,
            'in order to make it equal to zero, if
            'it is not in the break even point
            
            If current_cell > 0 Then
                holdback = holdback + current_cell
                current_cell = 0
            End If
            
    '        current_cell.Activate
        Next counter
           
        For counter = STARTING_FROM_COLUMN To cell_with_break_even.Column Step 1
            With tbl_output
                Set current_cell = .Cells(current_row, counter)
            End With
    
            If current_cell < 0 Then
                holdback = holdback + current_cell
                current_cell = 0
            Else
                If holdback + current_cell < 0 Then
                    holdback = holdback + current_cell
                    current_cell = 0
                Else
                    current_cell = current_cell + holdback
                    holdback = 0
                End If
            End If
        Next counter
    End If
    
    Set range_for_analysis = Nothing
    Set cell_with_break_even = Nothing
    Set cell = Nothing
    Set current_cell = Nothing
   
   On Error GoTo 0
   Exit Sub

RedAndBlackRecalculation_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RedAndBlackRecalculation of Modul mod_RedAndBlackRecalculation"
End Sub

Function last_column(Optional str_sheet As String, Optional row_to_check As Long = 1) As Long

    Dim shSheet  As Worksheet
    
    If str_sheet = vbNullString Then
        Set shSheet = ActiveSheet
    Else
        Set shSheet = Worksheets(str_sheet)
    End If
    
    last_column = shSheet.Cells(row_to_check, shSheet.Columns.Count).End(xlToLeft).Column

End Function

Public Function RGB2HTMLColor(B As Byte, G As Byte, R As Byte) As String

    Dim HexR As Variant, HexB As Variant, HexG As Variant
    Dim sTemp As String

    On Error GoTo ErrorHandler

    'R
    HexR = Hex(R)
    If Len(HexR) < 2 Then HexR = "0" & HexR

    'Get Green Hex
    HexG = Hex(G)
    If Len(HexG) < 2 Then HexG = "0" & HexG

    HexB = Hex(B)
    If Len(HexB) < 2 Then HexB = "0" & HexB

    RGB2HTMLColor = HexR & HexG & HexB
    Debug.Print "Enter RGB, without caring for the real colors, the function knows what it is doing."
    Debug.Print "IF 50D092 then &H0050D092&"

    Exit Function
    
ErrorHandler:
    Debug.Print "RGB2HTMLColor was not successful"
End Function

Public Sub OnStart()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Application.EnableEvents = False

End Sub

Public Sub OnEnd()
    
    'Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
End Sub

Public Function sum_range(my_range As Range) As Double
    
    Dim cell As Range
    
    sum_range = 0
    
    For Each cell In my_range
        sum_range = sum_range + cell.Value
    Next
    
End Function