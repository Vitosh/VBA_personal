Attribute VB_Name = "tddMain"
Option Explicit
Option Private Module

Sub Tdd(Optional export As Boolean = False)
    
    On Error Resume Next

    Dim specs           As New tddSpecSuite
    
    Debug.Print "Test report from " & Environ("Username") & vbCrLf & "START: " & Now() & vbCrLf
    PUB_STR_ERROR_REPORT = "Test report from " & Environ("Username") & vbCrLf & "START: " & Now() & vbCrLf
    
    '---------------------
    'Tests start here ---v
    specs.It("001").Expect(SumArray(Array(3, 3, 3))).ToEqual 9
    specs.It("002").Expect(SumArray(Array(-100, -100, -200, -200), 1)).ToEqual -400
    specs.It("003").Expect(SumArray(Array(-100, -100, -200, -200))).ToEqual -600
    
    Dim myArray As Variant
    myArray = Array(2, 3, 4)
    specs.It("004").Expect(IsArrayAllocated(Array())).ToEqual False
    specs.It("005").Expect(IsArrayAllocated(myArray)).ToEqual True
    
    specs.It("006").Expect(ColumnNumberToLetter(26)).ToEqual "Z"
    specs.It("007").Expect(ColumnNumberToLetter(27)).ToEqual "AA"
    
    specs.It("008").Expect(GetRgb(256)).ToEqual "R=0, G=1, B=0"
    specs.It("009").Expect(GetRgb(255)).ToEqual "R=255, G=0, B=0"
    specs.It("010").Expect(GetRgb(1)).ToNotEqual "R=2, G=0, B=0"
    specs.It("011").Expect(GetRgb(1)).ToEqual "R=1, G=0, B=0"
    
    '---------------------
    'Tests end here -----^
    
    tddSpecInlineRunner.RunSuite specs
    specs.TotalTests
    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & "END: " & Now() & vbCrLf
    Debug.Print "END: " & Now() & vbCrLf
    If export Then PrintToNotepad
    On Error GoTo 0
    
End Sub
