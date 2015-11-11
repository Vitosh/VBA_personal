Option Explicit

Private p_counter_value         As Long
Private p_flagged               As Variant

Public Sub IncrementCounter(Optional value As Long = 1)
    p_counter_value = p_counter_value + value
    ReDim Preserve p_flagged(UBound(p_flagged) + 1)
End Sub

Public Sub ResetCounter(Optional value As Long = 1)
    p_counter_value = value
    Call ResetFlagsArray
End Sub

Public Property Get Counter() As Long
    Counter = p_counter_value
End Property

Public Sub Flag()
    p_flagged(Counter) = True
End Sub

Public Sub UnFlag()
    p_flagged(Counter) = False
End Sub

Public Property Get IsFlagged() As Boolean
    IsFlagged = p_flagged(Counter)
End Property

Private Sub Class_Initialize()
    Call ResetFlagsArray
End Sub

Public Sub ResetFlagsArray()
    ReDim p_flagged(1)
End Sub

'TESTS ARE HERE:
Sub TDD()
    
    Dim specs               As New SpecSuite
    Dim test_calendar       As New cls_calendar
    Dim test_plan           As New cls_plan
    Dim test_counter        As New cls_counter
    
    test_counter.IncrementCounter
    test_counter.IncrementCounter
    specs.It("counter_c9").Expect(test_counter.Counter).ToEqual 2
    specs.It("counter_c10").Expect(test_counter.Counter).ToNotEqual 3
    test_counter.ResetCounter
    specs.It("counter_c11").Expect(test_counter.Counter).ToEqual 1
    test_counter.IncrementCounter (10)
    specs.It("counter_c12").Expect(test_counter.Counter).ToEqual 11
    
    test_counter.ResetCounter
    test_counter.IncrementCounter
    test_counter.Flag
    
    specs.It("counter_c13").Expect(test_counter.IsFlagged).ToEqual True
    test_counter.IncrementCounter
    specs.It("counter_c14").Expect(test_counter.IsFlagged).ToEqual False
    test_counter.Flag
    specs.It("counter_c14").Expect(test_counter.IsFlagged).ToEqual True
    test_counter.UnFlag
    specs.It("counter_c14").Expect(test_counter.IsFlagged).ToEqual False
    
    InlineRunner.RunSuite specs
    
    Set test_calendar = Nothing
    Set test_plan = Nothing
    Set specs = Nothing
    Set test_counter = Nothing
    
End Sub
