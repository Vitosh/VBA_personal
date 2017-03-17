Option Explicit

Sub Dictionaries()

    Dim l_counter1      As Long
    Dim l_counter2      As Long

    Dim dicts(7)        As Variant
    Dim predecessors    As Variant
    
    Dim node            As New Dictionary
    
    Set predecessors = New Dictionary
    
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
    
    For l_counter1 = 0 To UBound(dicts)
        
        Set node = dicts(l_counter1)
        If Not b_key_in_dict(predecessors, node.Keys(0)) Then
            Debug.Print node.Keys(0)
            predecessors.Add node.Keys(0), 0
        End If
        
        'Check if node has no children
        If UBound(node(node.Keys(0))) > 0 Then
            For l_counter2 = 0 To UBound(node.Items)
                If Not (b_key_in_dict(predecessors, node.Items(l_counter2)(0))) Then
                    predecessors.Add node.Items(l_counter2)(0), 0
                Else
                    predecessors.Item(node.Items(l_counter2)(0)) = (node.Items(l_counter2)(0)) + 1
                End If
            Next l_counter2
        End If
    Next l_counter1
    
'   Set k = dicts(5)
'   Debug.Print k.Item(3)(0)            'First Item in the array in k with key 3
'   Debug.Print k.Item(3)(1)            'Second Item in the array in k with key 3
'   Debug.Print UBound(k.Item(3))       'Size of items in the array in k with key 3 (-1)
'   Debug.Print k.Keys(0)               'First key of k
'   Debug.Print UBound(k.Keys)          'Size of keys in k (-1)
    
End Sub

Public Function b_key_in_dict(ByVal dict As Dictionary, ByVal key As String) As Boolean
'called like ->     b_key_in_dict(dicts(0),5)
' OR just use EXIST

    Dim l_counter       As Long
    
    b_key_in_dict = False
    For l_counter = 0 To UBound(dict.Keys)
        If dict.Keys(l_counter) = key Then b_key_in_dict = True
    Next l_counter

End Function

