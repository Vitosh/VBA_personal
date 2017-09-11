Option Explicit

Public Function get_last_day_of_month(ByVal my_date As Date) As Date
    get_last_day_of_month = DateSerial(Year(my_date), month(my_date) + 1, 0)
End Function

Public Function get_first_day_of_month(ByVal my_date As Date) As Date
    get_first_day_of_month = DateSerial(Year(my_date), month(my_date), 1)
End Function

Public Function add_months(ByVal my_date As Date, ByVal i_month As Long) As Date
    add_months = get_last_day_of_month(DateAdd("m", i_month, my_date))
End Function

Public Function add_months_and_get_first_date(ByVal my_date As Date, ByVal i_month As Long) As Date
    add_months_and_get_first_date = get_first_day_of_month(DateAdd("m", i_month, my_date))
End Function

Public Function date_diff_in_months(a As Date, b As Date) As Long
    date_diff_in_months = DateDiff("m", a, b)
End Function

' change the style
Public Function fnDatLastDayOfMonth(ByVal myDate As Date) As Date
    fnDatLastDayOfMonth = DateSerial(Year(myDate), Month(myDate) + 1, 0)
End Function

---------a bit better format:--------------

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

Public Function AddMonthsAndGetFirstDate(ByVal my_date As Date, ByVal i_month As Long) As Date
    AddMonthsAndGetFirstDate = GetFirstDayOfMonth(DateAdd("m", i_month, my_date))
End Function

Public Function DateDiffInMonths(a As Date, b As Date) As Long
    DateDiffInMonths = DateDiff("m", a, b)
End Function

Public Function DateLastDayOfMonth(ByVal myDate As Date) As Date
    DateLastDayOfMonth = DateSerial(Year(myDate), Month(myDate) + 1, 0)
End Function
