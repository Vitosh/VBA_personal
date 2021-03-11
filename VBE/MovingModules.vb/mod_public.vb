Option Explicit

Public Const L_STARTING_ROW = 6

Public Const L_RATE6_VERTRAG_COL = 6
Public Const L_RATE6_TERMIN_COL = 10
Public Const L_RATE5PR_VERTRAG_COL = 8
Public Const L_RATE5PR_TERMIN_COL = 12
Public Const STR_FERTIG = "Fertig!"
Public Const STR_SCHADENSERSATZ = "Schadensersatz"
Public Const L_FIRST_COLUMN_TO_WRITE = 21
Public Const L_ROW_WITH_DATES = 5

Public Const L_WOHNFLAECHE_COL = 15

Public obj_cal                      As cls_calendar

Public dbl_eur_m2                   As Double
Public dbl_eur_garage               As Double

Public l_counter                    As Long

Public r_range_4_dates              As Range
Public my_cell                      As Range
