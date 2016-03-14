Sub ExtendContentFromRight()
    
    Dim rng_first           As Range

    Set rng_first = Selection.Cells(1, 1)
    
    Selection.Formula = rng_first.Formula
    
    Set rng_first = Nothing
    
 End Sub
