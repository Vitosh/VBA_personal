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
    'Test Scenario #1
    TestMeSample
    Dim myarr(16) As Variant
    Dim arrCounter As Long
    Dim myCell As Range
    
    myarr(1) = 1.81859485365136
    myarr(2) = -4.79462137331569
    myarr(3) = -0.713935644387188
    myarr(4) = -8.38308001079428
    myarr(5) = 24.9643391023361
    myarr(6) = -27.4617351821139
    myarr(7) = 64.2321735505502
    myarr(8) = -88.9405995522673
    myarr(9) = -127.858501929498
    myarr(10) = 101.737867039937
    myarr(11) = 146.707455130634
    myarr(12) = -120.333197895024
    myarr(13) = 772.275323251858
    myarr(14) = 1129.5172126244
    myarr(15) = 1312.97247658607
    myarr(16) = -349.11864840751

    For Each myCell In tblInput.Range("A1:B8")
        Increment arrCounter
        specs.It("Scenario 1." & CStr(arrCounter)).Expect(myarr(arrCounter)).ToEqual myCell.value
    Next myCell
    
    'Test Scenario #2
    specs.It("Scenario 2.1").Expect(SumArray(Array(1, 2, 3))).ToEqual 6
    specs.It("Scenario 2.2").Expect(SumArray(Array(3, 3, 3))).ToEqual 9
    specs.It("Scenario 2.3").Expect(SumArray(Array(3, 4, 3))).ToNotEqual 9
    specs.It("Scenario 2.4").Expect(SumArray(Array(3, 3, 100), 1)).ToEqual 6
    specs.It("Scenario 2.5").Expect(SumArray(Array(3, 3, 100))).ToEqual 106
    specs.It("Scenario 2.6").Expect(SumArray(Array(-3, -3))).ToEqual -6
    
    'Tests Scenario #3
    specs.It("Scenario 3.1").Expect(ColumnNumberToLetter(26)).ToEqual "Z"
    specs.It("Scenario 3.2").Expect(ColumnNumberToLetter(1)).ToEqual "A"
    
    '---------------------
    'Tests end here -----^
    tddSpecInlineRunner.RunSuite specs
    specs.TotalTests
    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & "END: " & Now() & vbCrLf
    Debug.Print "END: " & Now() & vbCrLf
    If export Then PrintToNotepad
    On Error GoTo 0
    
End Sub

Public Sub MakeAllValues()
    
    Dim myCell As Range
    Dim i As Long
    Dim str As String
    
    For Each myCell In Selection
        Increment i
        str = vbTab & "myArr(" & i & ")= "
        
        If Len(myCell) > 0 Then
            If IsDate(myCell) Then
                str = str & "CDate(""" & myCell & """)"
            Else
                If Not IsNumeric(myCell) Then
                    str = str & """" & myCell & """"
                Else
                    str = str & ChangeCommas(myCell.value)
                End If
            End If
        Else
            If myCell.HasFormula Then
                str = str & """"""
            Else
                str = str & 0
            End If
        End If
        
        Debug.Print str
    Next myCell
    
End Sub

Sub TestMeSample()
    
    Dim myCell As Range
    Dim myVal As Variant
    
    For Each myCell In tblInput.Range("A1:B8")
        myVal = myVal * 1.5 + 2
        myCell = myVal * Sin(myVal)
    Next
    
End Sub

