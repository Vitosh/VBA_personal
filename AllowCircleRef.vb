'circle reference, circle ref
'Excel Options>Formulas

Sub RoundingCircle()

    With Application
        .Iteration = True
        .MaxIterations = 1
        .MaxChange = 0.0001
    End With
    
End Sub
