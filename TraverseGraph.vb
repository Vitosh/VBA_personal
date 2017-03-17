'Exercises: graph Algorithms
'This document defines the in-class exercises assignments for the "Algorithms" course @ Software University.
'For the following exercises you are given a Visual Studio solution "Graph-Algorithms-Lab" holding portions of the source code + unit tests. You can download it from the course's page.
'Part I - Traverse a Graph to Find Its Connected Components

Option Explicit

Public visited      As Variant
Public graph        As Variant

Public Sub mains()

    Dim l_counter       As Long
    Dim g1              As Variant
    Dim g2              As Variant
    Dim g3              As Variant
    Dim g4              As Variant
    Dim g5              As Variant
    Dim g6              As Variant
    Dim g7              As Variant
    Dim g8              As Variant
    Dim g9              As Variant
    
    g1 = Array(3, 6)
    g2 = Array(3, 4, 5, 6)
    g3 = Array(8)
    g4 = Array(0, 1, 5)
    g5 = Array(1, 6)
    g6 = Array(1, 3)
    g7 = Array(0, 1, 4)
    g8 = Array()
    g9 = Array(2)
    
    graph = Array(g1, g2, g3, g4, g5, g6, g7, g8, g9)
    
    ReDim visited(0)
    
    For l_counter = LBound(graph) To UBound(graph)
    
        If UBound(graph(l_counter)) >= 0 Then
            If Not b_value_in_array(graph(l_counter)(0), visited) Then
                Call DFS(graph(l_counter)(0))
                Debug.Print "---------------------"
            End If
        Else
            Debug.Print l_counter
            Debug.Print "---------------------"
        End If
    Next l_counter
End Sub

Public Sub DFS(ByVal str_node As String)
    
    Dim nodes       As Variant
    Dim cur_node    As String
    Dim child_node  As Variant
    Dim k           As Variant
    
    nodes = Array(0, str_node)
    ReDim Preserve visited(UBound(visited) + 1)
    visited(UBound(visited)) = str_node
    
    While UBound(nodes) > 0
        cur_node = nodes(UBound(nodes))
        Debug.Print cur_node
        
        ReDim Preserve nodes(UBound(nodes) - 1)
        
        child_node = graph(cur_node)
        
        For Each k In child_node
            
            If Not b_value_in_array(k, visited) Then
                ReDim Preserve nodes(UBound(nodes) + 1)
                nodes(UBound(nodes)) = k
                
                ReDim Preserve visited(UBound(visited) + 1)
                visited(UBound(visited)) = k
                
            End If
            
        Next k
    Wend
    
End Sub

Public Function b_value_in_array(my_value As Variant, my_array As Variant, Optional b_is_string As Boolean = False) As Boolean

    Dim l_counter   As Long

    If b_is_string Then
        my_array = Split(my_array, ":")
    End If

    For l_counter = LBound(my_array) To UBound(my_array)
        my_array(l_counter) = CStr(my_array(l_counter))
    Next l_counter

    b_value_in_array = Not IsError(Application.Match(CStr(my_value), my_array, 0))
    
End Function
