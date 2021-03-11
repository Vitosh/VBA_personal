Option Explicit
Option Private Module

Public Sub RunSuite(specs As SpecSuite, _
                    Optional ShowFailureDetails As Boolean = True, _
                    Optional ShowPassed As Boolean = False, _
                    Optional ShowSuiteDetails As Boolean = False)

    Dim SuiteCol            As New Collection

    SuiteCol.Add specs
    RunSuites SuiteCol, ShowFailureDetails, ShowPassed, ShowSuiteDetails

End Sub

Public Sub RunSuites(SuiteCol As Collection, _
                    Optional ShowFailureDetails As Boolean = True, _
                    Optional ShowPassed As Boolean = False, _
                    Optional ShowSuiteDetails As Boolean = True)

    Dim Suite               As SpecSuite
    Dim Spec                As SpecDefinition

    Dim TotalCount          As Long
    Dim FailedSpecs         As Long
    Dim PendingSpecs        As Long

    Dim ShowingResults      As Boolean
    Dim Indentation         As String

    For Each Suite In SuiteCol
        If Not Suite Is Nothing Then
            TotalCount = TotalCount + Suite.SpecsCol.Count

            For Each Spec In Suite.SpecsCol
                If Spec.result = SpecResult.FAIL Then
                    FailedSpecs = FailedSpecs + 1
                ElseIf Spec.result = SpecResult.Pending Then
                    PendingSpecs = PendingSpecs + 1
                End If
            Next Spec
        End If
    Next Suite

    Debug.Print "= " & SummaryMessage(TotalCount, FailedSpecs, PendingSpecs) & " = " & GetDateAndTime & " =========================" & vbCrLf
    str_error_report = str_error_report & "= " & SummaryMessage(TotalCount, FailedSpecs, PendingSpecs) & " = " & GetDateAndTime & " ========================="
    
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
                If Spec.result = SpecResult.FAIL Then
                    Debug.Print Indentation & FailureMessage(Spec, ShowFailureDetails, Indentation)
                    str_error_report = str_error_report & vbCrLf & Indentation & FailureMessage(Spec, ShowFailureDetails, Indentation)
                    ShowingResults = True
                ElseIf Spec.result = SpecResult.Pending Then
                    Debug.Print Indentation & PendingMessage(Spec)
                    str_error_report = str_error_report & vbCrLf & Indentation & PendingMessage(Spec)
                    ShowingResults = True
                ElseIf ShowPassed Then
                    Debug.Print Indentation & PassingMessage(Spec)
                    str_error_report = str_error_report & vbCrLf & Indentation & PassingMessage(Spec)
                    ShowingResults = True
                End If
            Next Spec
        End If
    Next Suite

    If ShowingResults Then
        Debug.Print "==="
        str_error_report = str_error_report & vbCrLf & "===" & vbCrLf
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

Private Function FailureMessage(Spec As SpecDefinition, ShowFailureDetails As Boolean, Indentation As String) As String

    Dim FailedExpectation   As SpecExpectation
    Dim i                   As Long
    
    FailureMessage = ResultMessage(Spec, "X")
    
    If ShowFailureDetails Then
        FailureMessage = FailureMessage & vbNewLine
        
        For Each FailedExpectation In Spec.FailedExpectations
            FailureMessage = FailureMessage & Indentation & "  " & FailedExpectation.FailureMessage
            
            If i + 1 <> Spec.FailedExpectations.Count Then FailureMessage = FailureMessage & vbNewLine
            i = i + 1
        Next FailedExpectation
    End If
    
End Function

Private Function PendingMessage(Spec As SpecDefinition) As String
    
    PendingMessage = ResultMessage(Spec, ".")
    
End Function

Private Function PassingMessage(Spec As SpecDefinition) As String

    PassingMessage = ResultMessage(Spec, "+")

End Function

Private Function ResultMessage(Spec As SpecDefinition, Symbol As String) As String

    ResultMessage = Symbol & " "

    If Spec.ID <> "" Then
        ResultMessage = ResultMessage & Spec.ID & ": "
    End If

    ResultMessage = ResultMessage & Spec.Description

End Function

Private Function SuiteMessage(Suite As SpecSuite) As String

    Dim HasFailures     As Boolean
    Dim Spec            As SpecDefinition

    For Each Spec In Suite.SpecsCol
        If Spec.result = SpecResult.FAIL Then
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
