Option Explicit

Private p_last_row              As Long
Private p_length_of_calendar    As Long
Private p_rightest_column       As Long

Private p_date_first_month      As Date
Private p_date_last_month       As Date

Private p_range_4_dates         As Range
'

Public Property Get Range4Dates() As Range
    Range4Dates = p_range_4_dates
End Property

Public Property Let Range4Dates(value As Range)
    p_range_4_dates = value
End Property

Public Property Get RightestColumn() As Long
    RightestColumn = p_rightest_column
End Property

Public Property Let RightestColumn(value As Long)
    p_rightest_column = value
End Property

Public Property Get CalendarLength() As Long
    CalendarLength = p_length_of_calendar
End Property

Public Property Let CalendarLength(value As Long)
    p_length_of_calendar = value
End Property

Public Property Get LastMonth() As Date
    LastMonth = p_date_last_month
End Property

Public Property Let LastMonth(value As Date)
    p_date_last_month = value
End Property

Public Property Get FirstMonth() As Date
    FirstMonth = p_date_first_month
End Property

Public Property Let FirstMonth(value As Date)
    p_date_first_month = value
End Property

Public Property Get LastRow() As Long
    LastRow = p_last_row
End Property

Public Property Let LastRow(value As Long)
    p_last_row = value
End Property
