Option Explicit

Public Function GetLastDayOfMonth(ByVal myDate As Date) As Date
    GetLastDayOfMonth = DateSerial(Year(myDate), Month(myDate) + 1, 0)
End Function

Public Function GetFirstDayOfMonth(ByVal myDate As Date) As Date
    GetFirstDayOfMonth = DateSerial(Year(myDate), Month(myDate), 1)
End Function

Public Function AddMonths(ByVal myDate As Date, ByVal lngMonth As Long) As Date
    AddMonths = GetLastDayOfMonth(DateAdd("m", lngMonth, myDate))
End Function

Public Function AddMonthsAndGetFirstDate(ByVal my_date As Date, ByVal lngMonth As Long) As Date
    AddMonthsAndGetFirstDate = GetFirstDayOfMonth(DateAdd("m", lngMonth, my_date))
End Function

Public Function DateDiffInMonths(a As Date, b As Date) As Long
    DateDiffInMonths = DateDiff("m", a, b)
End Function

Sub TestMe()

    Debug.Print GetLastDayOfMonth(DateSerial(2020, 2, 22))
    Debug.Print GetLastDayOfMonth(DateSerial(2021, 2, 22))
    
    Debug.Print GetFirstDayOfMonth(DateSerial(2021, 2, 22))
    Debug.Print AddMonths(DateSerial(2020, 2, 23), 3)
    Debug.Print AddMonthsAndGetFirstDate(DateSerial(2020, 2, 23), 3)
    
    Debug.Print DateDiffInMonths(DateSerial(1988, 8, 18), DateSerial(1998, 10, 18))
    
End Sub
