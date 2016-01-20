Call FormatDin(my_cell)
Call FormatDark(my_cell)

Public Sub FormatDark(ByRef my_cell As range)
  my_cell.Interior.ThemeColor = xlThemeColorDark1
  my_cell.Interior.TintAndShade = -0.249946592608417
End Sub

Public Sub FormatDin(ByRef my_cell As range)
  my_cell.Font.Name = "DIN-Light"
End Sub
