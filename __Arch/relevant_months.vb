Option Explicit

Public Function relevant_month(ByVal dt_date As Date) As String
    
    relevant_month = WorksheetFunction.Choose(Month(dt_date), "jan", "feb", ",mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    relevant_month = relevant_month & "_" & Right(Year(dt_date), 2)

End Function

Public Function relevant_month_de(ByVal dt_date As Date) As String

    relevant_month_de = LCase(MonthName(Month(dt_date), True) & "_" & Right(Year(dt_date), 2))

End Function

Public Sub CheckName()
    
    Debug.Print relevant_month_de(Now() + 40)
    Debug.Print relevant_month(Now() + 40)
    
End Sub

Public Function bad_example()

If public_date <= #12/31/2005# Then
relevant_month = "dec_05"

ElseIf (public_date > #12/31/2005# And public_date <= #1/31/2006#) Then
relevant_month = "jan_06"
'300 lines more with elseifs
Else
relevant_month = "jan_18"
End If

send_relevant_month = relevant_month

End Function

  Does not compile,
  Does not take lump years,
  Does not run automatically
