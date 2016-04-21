Option Explicit

Public Sub NumbersAndDictionary()

    Dim l_counter_0     As Long
    Dim l_counter_1     As Long
    Dim l_counter_2     As Long
    Dim l_value         As Long
    
    Dim my_dict         As Object
    Dim my_str          As String
    
    Set my_dict = CreateObject("Scripting.Dictionary")
    
    For l_counter_0 = 65 To 76
        my_str = CStr(l_counter_0)
        For l_counter_1 = 1 To Len(my_str)
            l_value = CLng(Mid(my_str, l_counter_1, 1))
            
            If Not my_dict.Exists(l_value) Then
                my_dict.Add l_value, 1
            Else
                my_dict(l_value) = my_dict(l_value) + 1
            End If
            
        Next l_counter_1
    Next l_counter_0
    
    Call PrintDictionary(my_dict)
    
    Set my_dict = Nothing
    
End Sub

Public Sub PrintDictionary(my_dict As Object)
    
    Dim l_counter_0     As Long

    For l_counter_0 = 0 To my_dict.Count - 1
        Debug.Print l_counter_0; " "; my_dict.Item(l_counter_0)
    Next l_counter_0
End Sub
