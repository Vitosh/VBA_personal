Sub TDD()
    
    Dim specs               As New SpecSuite
    Dim test_calendar       As New cls_calendar
    Dim test_plan           As New cls_plan
    Dim test_counter        As New cls_counter
    
    test_calendar.IncrementRow
    specs.It("cls_c1").Expect(test_calendar.CurrentRow).ToEqual 1
    
    test_calendar.IncrementRow
    test_calendar.CurrentRow = 5
    test_calendar.IncrementRow
    specs.It("cls_c2").Expect(test_calendar.CurrentRow).ToEqual 6
    
    test_calendar.LeftDate = "01.09.2015"
    test_calendar.RightDate = "01.08.2021"
    specs.It("cls_c3").Expect(test_calendar.Duration).ToEqual 71
    
    test_plan.LastLines_Row = 30
    test_plan.LastLines_Row = 31
    test_plan.LastLines_Row = 1000
    specs.It("plan_c4").Expect(test_plan.LastLines_Row(1)).ToEqual 30
    specs.It("plan_c5").Expect(test_plan.LastLines_Row(2)).ToEqual 31
    specs.It("plan_c6").Expect(test_plan.LastLines_Row(3)).ToEqual 1000
    
    specs.It("plan_c7").Expect(test_plan.LastLines_Row_Count).ToEqual 3
    specs.It("plan_c8").Expect(test_plan.LastLines_Row_Count).ToNotEqual 4
    
    test_counter.IncrementCounter
    test_counter.IncrementCounter
    specs.It("counter_c9").Expect(test_counter.Counter).ToEqual 2
    specs.It("counter_c10").Expect(test_counter.Counter).ToNotEqual 3
    
    test_counter.ResetCounter
    specs.It("counter_c11").Expect(test_counter.Counter).ToEqual 0
    
    test_counter.IncrementCounter (10)
    specs.It("counter_c12").Expect(test_counter.Counter).ToEqual 10
    
    test_counter.IncrementCounter
    specs.It("counter_c12a").Expect(test_counter.Counter).ToEqual 11
    
    test_counter.DecrementCounter
    specs.It("counter_c12b").Expect(test_counter.Counter).ToEqual 10
    
    test_counter.ResetCounter
    test_counter.DecrementCounter
    specs.It("counter_c12c").Expect(test_counter.Counter).ToEqual -1

    test_counter.ResetCounter
    test_counter.IncrementCounter
    test_counter.Flag
    specs.It("counter_c13").Expect(test_counter.IsFlagged).ToEqual True
    
    test_counter.IncrementCounter
    specs.It("counter_c14").Expect(test_counter.IsFlagged).ToEqual False
    
    test_counter.Flag
    specs.It("counter_c15").Expect(test_counter.IsFlagged).ToEqual True
    
    test_counter.UnFlag
    specs.It("counter_c16").Expect(test_counter.IsFlagged).ToEqual False
    
    InlineRunner.RunSuite specs
    
    Set test_calendar = Nothing
    Set test_plan = Nothing
    Set specs = Nothing
    Set test_counter = Nothing
    
End Sub

'Unit tests are here:
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
