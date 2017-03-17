Sub GetUsersUsingWorkbook()

    Dim users           As Variant
    Dim l_counter       As Long
    
    users = ActiveWorkbook.UserStatus
    Debug.Print "Total Users using the current WorkBook: " & UBound(users)
    
    For l_counter = 1 To UBound(users)
        Debug.Print
        Debug.Print users(l_counter, 1)
        Debug.Print users(l_counter, 2)
        Debug.Print users(l_counter, 3)
    Next l_counter

End Sub
