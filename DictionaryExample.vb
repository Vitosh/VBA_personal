Sub MyDictionary()
    
    'Add
    Dim myDict As New Scripting.Dictionary
    myDict.Add "Peter", "Peter is a friend."
    myDict.Add "George", "George is a guy I know."
    myDict.Add "Salary", 1000
    
    'Exists
    If myDict.Exists("Salary") Then
        Debug.Print myDict("Salary")
        myDict("Salary") = myDict("Salary") * 2
        Debug.Print myDict("Salary")
    End If
    
    'Remove
    If myDict.Exists("George") Then
        myDict.Remove ("George")
    End If
    
    'Items
    Dim item As Variant
    For Each item In myDict.Items
        Debug.Print item
    Next item
        
    'Keys
    Dim key As Variant
    For Each key In myDict.Keys
        Debug.Print key
    Next key
    
    'Remove All
    myDict.RemoveAll
    
    'Compare Mode
    myDict.CompareMode = BinaryCompare
    
    myDict.Add "PeTeR", "Peter written as PeTeR"
    myDict.Add "PETeR", "Peter written as PETeR"
    PrintDictionary myDict
    
End Sub


Public Sub PrintDictionary(myDict As Object)
    
    Dim key     As Variant
    For Each key In myDict.Keys
        Debug.Print key; "-->"; myDict(key)
    Next key
    
End Sub
