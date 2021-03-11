Public Enum ExpectResult

    PASS
    FAIL

End Enum

Public Actual                       As Variant
Public Expected                     As Variant
Public result                       As ExpectResult
Public FailureMessage               As String

Public Sub ToEqual(Expected As Variant)
    
    check IsEqual(Me.Actual, Expected), "to equal", Expected:=Expected

End Sub
Public Sub ToNotEqual(Expected As Variant)

    check IsEqual(Me.Actual, Expected), "to not equal", Expected:=Expected, Inverse:=True
    
End Sub

Private Function IsEqual(Actual As Variant, Expected As Variant) As Variant
    
    Dim l_count         As Long
    
    'here added additional value
    If IsArray(Expected) Then
        If UBound(Expected) <> UBound(Actual) Then IsEqual = False: Exit Function
        
        For l_count = LBound(Expected) To UBound(Expected)
            If Not Expected(l_count) = Actual(l_count) Then IsEqual = False: Exit Function
        Next l_count
        
        IsEqual = True
        
    End If
    'end of additional value

    If IsError(Actual) Or IsError(Expected) Then
        IsEqual = False

    ElseIf IsObject(Actual) Or IsObject(Expected) Then
        IsEqual = "Unsupported: Can't compare objects"
    ElseIf VarType(Actual) = vbDouble And VarType(Expected) = vbDouble Then
        IsEqual = IsCloseTo(Actual, Expected, 15)
    Else
        IsEqual = Actual = Expected
    End If
    
End Function

Public Sub ToBeDefined()
    Debug.Print "Excel-TDD: DEPRECATED, ToBeDefined() has been deprecated in favor of ToNotBeUndefined and will be removed in Excel-TDD v2.0.0"
    check IsUndefined(Me.Actual), "to be defined", Inverse:=True
End Sub

Public Sub ToBeUndefined()
    check IsUndefined(Me.Actual), "to be undefined"
End Sub

Public Sub ToNotBeUndefined()
    check IsUndefined(Me.Actual), "to not be undefined", Inverse:=True
End Sub

Private Function IsUndefined(Actual As Variant) As Variant
    IsUndefined = IsNothing(Actual) Or IsEmpty(Actual) Or IsNull(Actual) Or IsMissing(Actual)
End Function

Public Sub ToBeNothing()
    check IsNothing(Me.Actual), "to be nothing"
End Sub
Public Sub ToNotBeNothing()
    check IsNothing(Me.Actual), "to not be nothing", Inverse:=True
End Sub

Private Function IsNothing(Actual As Variant) As Variant
    If IsObject(Actual) Then
        If Actual Is Nothing Then
            IsNothing = True
        Else
            IsNothing = False
        End If
    Else
        IsNothing = False
    End If
End Function

Public Sub ToBeEmpty()
    check IsEmpty(Me.Actual), "to be empty"
End Sub

Public Sub ToNotBeEmpty()
    check IsEmpty(Me.Actual), "to not be empty", Inverse:=True
End Sub

Public Sub ToBeNull()
    check IsNull(Me.Actual), "to be null"
End Sub

Public Sub ToNotBeNull()
    check IsNull(Me.Actual), "to not be null", Inverse:=True
End Sub

Public Sub ToBeMissing()
    check IsMissing(Me.Actual), "to be missing"
End Sub

Public Sub ToNotBeMissing()
    check IsMissing(Me.Actual), "to not be missing", Inverse:=True
End Sub

Public Sub ToBeLessThan(Expected As Variant)
    check IsLT(Me.Actual, Expected), "to be less than", Expected:=Expected
End Sub

Public Sub ToBeLT(Expected As Variant)
    ToBeLessThan Expected
End Sub

Private Function IsLT(Actual As Variant, Expected As Variant) As Variant
    If IsError(Actual) Or IsError(Expected) Or Actual >= Expected Then
        IsLT = False
    Else
        IsLT = True
    End If
End Function

Public Sub ToBeLessThanOrEqualTo(Expected As Variant)
    check IsLTE(Me.Actual, Expected), "to be less than or equal to", Expected:=Expected
End Sub

Public Sub ToBeLTE(Expected As Variant)
    ToBeLessThanOrEqualTo Expected
End Sub

Private Function IsLTE(Actual As Variant, Expected As Variant) As Variant
    If IsError(Actual) Or IsError(Expected) Or Actual > Expected Then
        IsLTE = False
    Else
        IsLTE = True
    End If
End Function

Public Sub ToBeGreaterThan(Expected As Variant)
    check IsGT(Me.Actual, Expected), "to be greater than", Expected:=Expected
End Sub

Public Sub ToBeGT(Expected As Variant)
    ToBeGreaterThan Expected
End Sub

Private Function IsGT(Actual As Variant, Expected As Variant) As Variant
    If IsError(Actual) Or IsError(Expected) Or Actual <= Expected Then
        IsGT = False
    Else
        IsGT = True
    End If
End Function

Public Sub ToBeGreaterThanOrEqualTo(Expected As Variant)
    check IsGTE(Me.Actual, Expected), "to be greater than or equal to", Expected:=Expected
End Sub

Public Sub ToBeGTE(Expected As Variant)
    ToBeGreaterThanOrEqualTo Expected
End Sub

Private Function IsGTE(Actual As Variant, Expected As Variant) As Variant
    If IsError(Actual) Or IsError(Expected) Or Actual < Expected Then
        IsGTE = False
    Else
        IsGTE = True
    End If
End Function

Public Sub ToBeCloseTo(Expected As Variant, SignificantFigures As Long)
    check IsCloseTo(Me.Actual, Expected, SignificantFigures), "to be close to", Expected:=Expected
End Sub

Public Sub ToNotBeCloseTo(Expected As Variant, SignificantFigures As Long)
    check IsCloseTo(Me.Actual, Expected, SignificantFigures), "to be close to", Expected:=Expected, Inverse:=True
End Sub

Private Function IsCloseTo(Actual As Variant, Expected As Variant, SignificantFigures As Long) As Variant
    Dim ActualAsString As String
    Dim ExpectedAsString As String
    
    If SignificantFigures < 1 Or SignificantFigures > 15 Then
        IsCloseTo = "ToBeCloseTo/ToNotBeClose to can only compare from 1 to 15 significant figures"""
    ElseIf Not IsError(Actual) And Not IsError(Expected) Then
        ' Convert values to scientific notation strings and then compare strings
        If Actual > 1 Then
            ActualAsString = VBA.Format$(Actual, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e+0")
        Else
            ActualAsString = VBA.Format$(Actual, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e-0")
        End If

        If Expected > 1 Then
            ExpectedAsString = VBA.Format$(Expected, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e+0")
        Else
            ExpectedAsString = VBA.Format$(Expected, VBA.Left$("0.00000000000000", SignificantFigures + 1) & "e-0")
        End If
        
        IsCloseTo = ActualAsString = ExpectedAsString
    End If
End Function

Public Sub ToContain(Expected As Variant, Optional MatchCase As Boolean = True)
    If VarType(Me.Actual) = vbString Then
        Debug.Print "Excel-TDD: DEPRECATED ToContain has been changed to ToMatch in Excel-TDD v2.0.0"
        If MatchCase Then
            check Matches(Me.Actual, Expected), "to match", Expected:=Expected
        Else
            check Matches(VBA.UCase$(Me.Actual), VBA.UCase$(Expected)), "to match", Expected:=Expected
        End If
    Else
        check Contains(Me.Actual, Expected), "to contain", Expected:=Expected
    End If
End Sub

Public Sub ToNotContain(Expected As Variant, Optional MatchCase As Boolean = True)
    If VarType(Me.Actual) = vbString Then
        Debug.Print "Excel-TDD: DEPRECATED ToNotContain has been changed to ToMatch in Excel-TDD v2.0.0"
        If MatchCase Then
            check Matches(Me.Actual, Expected), "to not match", Expected:=Expected, Inverse:=True
        Else
            check Matches(VBA.UCase$(Me.Actual), VBA.UCase$(Expected)), "to not match", Expected:=Expected, Inverse:=True
        End If
    Else
        check Contains(Me.Actual, Expected), "to not contain", Expected:=Expected, Inverse:=True
    End If
End Sub

Private Function Contains(Actual As Variant, Expected As Variant) As Variant
    
    Dim i As Long
    
    If Not IsArray(Actual) Then
        Contains = "Error: Actual needs to be an Array or Collection for ToContain/ToNotContain"
    Else
        If TypeOf Actual Is Collection Then
            For i = 1 To Actual.Count
                If Actual.item(i) = Expected Then
                    Contains = True
                    Exit Function
                End If
            Next i
        Else
            For i = LBound(Actual) To UBound(Actual)
                If Actual(i) = Expected Then
                    Contains = True
                    Exit Function
                End If
            Next i
        End If
    End If
    
End Function

Public Sub ToMatch(Expected As Variant)

    check Matches(Me.Actual, Expected), "to match", Expected:=Expected

End Sub

Public Sub ToNotMatch(Expected As Variant)

    check Matches(Me.Actual, Expected), "to not match", Expected:=Expected, Inverse:=True

End Sub

Private Function Matches(Actual As Variant, Expected As Variant) As Variant
    If InStr(Actual, Expected) > 0 Then
        Matches = True
    Else
        Matches = False
    End If
End Function

Public Sub RunMatcher(Name As String, Message As String, ParamArray Arguments())

    Dim Expected        As String
    Dim i               As Long
    Dim HasArguments    As Boolean
        
    HasArguments = UBound(Arguments) >= 0
    For i = LBound(Arguments) To UBound(Arguments)
        If Expected = "" Then
            Expected = GetStringForValue(Arguments(i))
        ElseIf i = UBound(Arguments) Then
            If (UBound(Arguments) > 1) Then
                Expected = Expected & ", and " & GetStringForValue(Arguments(i))
            Else
                Expected = Expected & " and " & GetStringForValue(Arguments(i))
            End If
        Else
            Expected = Expected & ", " & GetStringForValue(Arguments(i))
        End If
    Next i
    
    If HasArguments Then
        check Application.Run(Name, Me.Actual, Arguments), Message, Expected:=Expected
    Else
        check Application.Run(Name, Me.Actual), Message
    End If

End Sub

Private Sub check(result As Variant, Message As String, Optional Expected As Variant, Optional Inverse As Boolean = False)

    If Not IsMissing(Expected) Then
        If IsObject(Expected) Then
            Set Me.Expected = Expected
        Else
            Me.Expected = Expected
        End If
    End If

    If VarType(result) = vbString Then
        Fails CStr(result)
    Else
        If Inverse Then
            result = Not result
        End If
        
        If result Then
            Passes
        Else
            Fails CreateFailureMessage(Message, Expected)
        End If
    End If
End Sub

Private Sub Passes()

    Me.result = ExpectResult.PASS

End Sub

Private Sub Fails(Message As String)

    Me.result = ExpectResult.FAIL
    Me.FailureMessage = Message

End Sub

Private Function CreateFailureMessage(Message As String, Optional Expected As Variant) As String
    CreateFailureMessage = "Expected " & GetStringForValue(Me.Actual) & " " & Message
    If Not IsMissing(Expected) Then
        CreateFailureMessage = CreateFailureMessage & " " & GetStringForValue(Expected)
    End If
End Function

Private Function GetStringForValue(value As Variant) As String

    If IsObject(value) Then
    
        If value Is Nothing Then
            GetStringForValue = "(Nothing)"
        Else
            GetStringForValue = "(Object)"
        End If
        
    ElseIf IsArray(value) Then
        GetStringForValue = "(Array)"
        
    ElseIf IsEmpty(value) Then
        GetStringForValue = "(Empty)"
        
    ElseIf IsNull(value) Then
        GetStringForValue = "(Null)"
        
    ElseIf IsMissing(value) Then
        GetStringForValue = "(Missing)"
        
    Else
        GetStringForValue = CStr(value)
        
    End If
    
    If GetStringForValue = "" Then
        GetStringForValue = "(Undefined)"
    End If
    
End Function

Private Function IsArray(value As Variant) As Boolean

    If Not IsEmpty(value) Then
        If IsObject(value) Then
            If TypeOf value Is Collection Then
                IsArray = True
            End If
        ElseIf VarType(value) = vbArray Or VarType(value) = 8204 Then
            IsArray = True
        End If
    End If

End Function
