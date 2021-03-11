Private pExpectations           As Collection
Private pFailedExpectations     As Collection

Public Enum SpecResult
    PASS
    FAIL
    Pending
End Enum

Public Description As String
Public ID As String

Public Property Get Expectations() As Collection
    
    If pExpectations Is Nothing Then
        Set pExpectations = New Collection
    End If
    
    Set Expectations = pExpectations

End Property
Private Property Let Expectations(value As Collection)
    
    Set pExpectations = value

End Property

Public Property Get FailedExpectations() As Collection

    If pFailedExpectations Is Nothing Then
        Set pFailedExpectations = New Collection
    End If
    
    Set FailedExpectations = pFailedExpectations
    
End Property
Private Property Let FailedExpectations(value As Collection)

    Set pFailedExpectations = value
    
End Property

Public Function Expect(Optional value As Variant) As SpecExpectation

    Dim Exp As New SpecExpectation
    
    If VarType(value) = vbObject Then
        Set Exp.Actual = value
    Else
        Exp.Actual = value
    End If
    Me.Expectations.Add Exp
    
    Set Expect = Exp
    
End Function

Public Function result() As SpecResult

    Dim Exp As SpecExpectation
    
    ' Reset failed expectations
    FailedExpectations = New Collection
    
    ' If no expectations have been defined, return pending
    If Me.Expectations.Count < 1 Then
        result = Pending
    Else
        ' Loop through all expectations
        For Each Exp In Me.Expectations
            ' If expectation fails, store it
            If Exp.result = FAIL Then
                FailedExpectations.Add Exp
            End If
        Next Exp
        
        ' If no expectations failed, spec passes
        If Me.FailedExpectations.Count > 0 Then
            result = FAIL
        Else
            result = PASS
        End If
    End If
    
End Function

Public Function ResultName() As String

    Select Case Me.result
        Case PASS: ResultName = "Pass"
        Case FAIL: ResultName = "Fail"
        Case Pending: ResultName = "Pending"
    End Select
    
End Function
