VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tddSpecDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pExpectations As Collection
Private pFailedExpectations As Collection
Public Description As String
Public Id As String

Public Enum SpecResult
    Pass
    Fail
    Pending
End Enum

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

Public Function Expect(Optional value As Variant) As tddSpecExpectation

    Dim Exp As New tddSpecExpectation
    
    If VarType(value) = vbObject Then
        Set Exp.Actual = value
    Else
        Exp.Actual = value
    End If
    Me.Expectations.Add Exp
    
    Set Expect = Exp
    
End Function

Public Function Result() As SpecResult

    Dim Exp As tddSpecExpectation
    
    FailedExpectations = New Collection
    If Me.Expectations.Count < 1 Then
        Result = Pending
    Else
        For Each Exp In Me.Expectations
             If Exp.Result = Fail Then
                FailedExpectations.Add Exp
            End If
        Next Exp
        
        If Me.FailedExpectations.Count > 0 Then
            Result = Fail
        Else
            Result = Pass
        End If
    End If
    
End Function

Public Function ResultName() As String
    
    Select Case Me.Result
        Case Pass:
            ResultName = "Pass"
        Case Fail:
            ResultName = "Fail"
        Case Pending:
            ResultName = "Pending"
    End Select
    
End Function

