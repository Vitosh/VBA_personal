Sub StartingTimer(ByRef myTime As Double)

    Debug.Print "Strating at:"
    Debug.Print Time
    myTime = Timer
    
End Sub

Sub EndingTimer(ByRef myTime As Double)

    Debug.Print "Ending at:"
    Debug.Print Time
    Debug.Print "Total time:"
    Debug.Print Format((Timer - myTime) / 86400, "hh:mm:ss")

End Sub

Sub TestAll()
    
    Dim myTime As Double
    StartingTimer myTime
    
    Stop    'PUT THE STUFF HERE!
    
    EndingTimer myTime
    
End Sub
