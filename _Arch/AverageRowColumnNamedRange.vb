Public Function calculate_avg_row(rng As Range, Optional l_row As Long = 1) As Double

    Dim my_start    As Range
    Dim my_end      As Range

    Set my_start = Cells(rng.Cells(l_row, 1).Row, rng.Cells(l_row, 1).Column)
    Set my_end = rng.Cells(l_row, rng.Columns.Count)

    Debug.Print my_start.Address
    Debug.Print my_end.Address

    calculate_avg_row = WorksheetFunction.Average(Range(my_start, my_end))

End Function

Option Explicit

Public Function calculate_avg(rng As Range, Optional l_starting_col As Long = 1, Optional l_end_col As Long = 1) As Double

    Dim my_start    As Range
    Dim my_end      As Range

    Set my_start = Cells(rng.Cells(1, 1).Row, l_starting_col + rng.Cells(1, 1).Column - 1)
    Set my_end = Cells(rng.Cells(rng.Rows.Count, l_end_col).Row, rng.Columns.Count - rng.Cells(1, l_end_col).Column + l_end_col)

    'Debug.Print my_start.Address
    'Debug.Print my_end.Address

    calculate_avg = WorksheetFunction.Average(Range(my_start, my_end))

End Function
