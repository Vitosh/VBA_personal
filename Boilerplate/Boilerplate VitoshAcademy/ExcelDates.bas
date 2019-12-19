Attribute VB_Name = "ExcelDates"
Option Explicit
Option Private Module

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
