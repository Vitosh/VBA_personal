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

