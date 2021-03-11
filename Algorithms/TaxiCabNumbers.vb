'https://en.wikipedia.org/wiki/Taxicab_number

Option Explicit

Public Sub TaxiCabNumber()
    
    Dim a           As Long
    Dim b           As Long
    Dim lastNumber  As Long
    Dim cnt         As Long
    
    lastNumber = 200
    
    Dim arrList     As Object
    Set arrList = CreateObject("System.Collections.ArrayList")

    For a = 1 To lastNumber
        For b = a + 1 To lastNumber
            
            Dim current As String
            current = a ^ 3 + b ^ 3
            
            'Debug.Assert (a <> 1 Or b <> 12) And (a <> 9 Or b <> 10)
            
            If arrList.contains(current) Then
                Debug.Print current
            Else
                arrList.Add (current)
            End If
            
            cnt = cnt + 1
        Next b
    Next a
    
End Sub
