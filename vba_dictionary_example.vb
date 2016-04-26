'vba dictionary
'vb dictionary
'vb dictionary example


Option Explicit

Sub Dictionaries()

    Dim dicts(7) As Variant
    
    Set dicts(0) = New Dictionary
    dicts(0).Add 5, Array(11)
    
    Set dicts(1) = New Dictionary
    dicts(1).Add 7, Array(11, 8)
    
    Set dicts(2) = New Dictionary
    dicts(2).Add 8, Array(9)
    
    Set dicts(3) = New Dictionary
    dicts(3).Add 11, Array(9, 10, 2)
    
    Set dicts(4) = New Dictionary
    dicts(4).Add 9, Array()
    
    Set dicts(5) = New Dictionary
    dicts(5).Add 3, Array(8, 10)
    
    Set dicts(6) = New Dictionary
    dicts(6).Add 2, Array()
    
    Set dicts(7) = New Dictionary
    dicts(7).Add 10, Array()
    
    Dim k   As New Dictionary
    Set k = dicts(5)
    
    Debug.Print k.Item(3)(0)            'First Item in the array in k with key 3
    Debug.Print k.Item(3)(1)            'Second Item in the array in k with key 3
    Debug.Print UBound(k.Item(3))       'Size of items in the array in k with key 3 (-1)
    Debug.Print k.Keys(0)               'First key of k
    Debug.Print UBound(k.Keys)          'Size of keys in k (-1)
    
End Sub
