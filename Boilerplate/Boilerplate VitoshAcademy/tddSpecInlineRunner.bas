Attribute VB_Name = "tddSpecInlineRunner"
Option Explicit
Option Private Module

Public Sub RunSuite(specs As tddSpecSuite, Optional ShowFailureDetails As Boolean = True, Optional ShowPassed As Boolean = False, Optional ShowSuiteDetails As Boolean = False)
    
    Dim SuiteCol As New Collection
    
    SuiteCol.Add specs
    RunSuites SuiteCol, ShowFailureDetails, ShowPassed, ShowSuiteDetails

End Sub

Public Sub RunSuites(SuiteCol As Collection, Optional ShowFailureDetails As Boolean = True, Optional ShowPassed As Boolean = False, Optional ShowSuiteDetails As Boolean = True)
    
    Dim Suite           As tddSpecSuite
    Dim Spec            As tddSpecDefinition
    Dim TotalCount      As Long
    Dim FailedSpecs     As Long
    Dim PendingSpecs    As Long
    Dim ShowingResults  As Boolean
    Dim Indentation     As String
    
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            TotalCount = TotalCount + Suite.SpecsCol.Count

            For Each Spec In Suite.SpecsCol
                If Spec.Result = SpecResult.Fail Then
                    FailedSpecs = FailedSpecs + 1
                ElseIf Spec.Result = SpecResult.Pending Then
                    PendingSpecs = PendingSpecs + 1
                End If
            Next Spec
        End If
    Next Suite
    
    Debug.Print "= " & SummaryMessage(TotalCount, FailedSpecs, PendingSpecs) & " = " & Now & " ========================="
    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & "= " & SummaryMessage(TotalCount, FailedSpecs, PendingSpecs) & " = " & Now & " =========================" & vbCrLf
    
    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            If ShowSuiteDetails Then
                Debug.Print SuiteMessage(Suite)
                Indentation = "  "
                ShowingResults = True
            Else
                Indentation = ""
            End If
        
            For Each Spec In Suite.SpecsCol
                If Spec.Result = SpecResult.Fail Then
                    Debug.Print Indentation & FailureMessage(Spec, ShowFailureDetails, Indentation)
                    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & Indentation & FailureMessage(Spec, ShowFailureDetails, Indentation) & vbCrLf
                    ShowingResults = True
                ElseIf Spec.Result = SpecResult.Pending Then
                    Debug.Print Indentation & PendingMessage(Spec)
                    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & Indentation & PendingMessage(Spec) & vbCrLf
                    ShowingResults = True
                ElseIf ShowPassed Then
                    Debug.Print Indentation & PassingMessage(Spec)
                    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & Indentation & PassingMessage(Spec) & vbCrLf
                    ShowingResults = True
                End If
            Next Spec
        End If
    Next Suite
    
    If ShowingResults Then
        Debug.Print "==="
        PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & "===" & vbCrLf
    End If
    
End Sub

Private Function SummaryMessage(TotalCount As Long, FailedSpecs As Long, PendingSpecs As Long) As String
    
    If FailedSpecs = 0 Then
        SummaryMessage = "PASS (" & TotalCount - PendingSpecs & " of " & TotalCount & " passed"
    Else
        SummaryMessage = "FAIL (" & FailedSpecs & " of " & TotalCount & " failed"
    End If
    
    If PendingSpecs = 0 Then
        SummaryMessage = SummaryMessage & ")"
    Else
        SummaryMessage = SummaryMessage & ", " & PendingSpecs & " pending)"
    End If
    
End Function

Private Function FailureMessage(Spec As tddSpecDefinition, ShowFailureDetails As Boolean, Indentation As String) As String

    Dim FailedExpectation As tddSpecExpectation
    Dim i As Long
    
    FailureMessage = ResultMessage(Spec, "X")
    
    If ShowFailureDetails Then
        FailureMessage = FailureMessage & vbNewLine
        
        For Each FailedExpectation In Spec.FailedExpectations
            FailureMessage = FailureMessage & Indentation & "  " & FailedExpectation.FailureMessage
            
            If i + 1 <> Spec.FailedExpectations.Count Then: FailureMessage = FailureMessage & vbNewLine
            i = i + 1
        Next FailedExpectation
    End If
    
End Function

Private Function PendingMessage(Spec As tddSpecDefinition) As String
    PendingMessage = ResultMessage(Spec, ".")
End Function

Private Function PassingMessage(Spec As tddSpecDefinition) As String
    PassingMessage = ResultMessage(Spec, "+")
End Function

Private Function ResultMessage(Spec As tddSpecDefinition, Symbol As String) As String
    ResultMessage = Symbol & " "
    
    If Spec.Id <> "" Then
        ResultMessage = ResultMessage & Spec.Id & ": "
    End If
    
    ResultMessage = ResultMessage & Spec.Description
End Function

Private Function SuiteMessage(Suite As tddSpecSuite) As String
    Dim HasFailures As Boolean
    Dim Spec As tddSpecDefinition
    
    For Each Spec In Suite.SpecsCol
        If Spec.Result = SpecResult.Fail Then
            HasFailures = True
            Exit For
        End If
    Next Spec
    
    If HasFailures Then
        SuiteMessage = "X "
    Else
        SuiteMessage = "+ "
    End If
    
    If Suite.Description <> "" Then
        SuiteMessage = SuiteMessage & Suite.Description
    Else
        SuiteMessage = SuiteMessage & Suite.SpecsCol.Count & " specs"
    End If
End Function

