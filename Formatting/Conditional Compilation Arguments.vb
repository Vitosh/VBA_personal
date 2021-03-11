'Conditional Compilation Arguments in Access
'To set them this is the code:

Application.SetOption "Conditional Compilation Arguments","A=4:B=10"
'To get them:

Application.GetOption("Conditional Compilation Arguments")
'They are printed like this: A = 4 : B = 10

'That is how to test it:

Sub TestMe()

    #If A = 1 Then
        Debug.Print "a is 1"
    #Else
        Debug.Print "a is not 1"
    #End If

End Sub
