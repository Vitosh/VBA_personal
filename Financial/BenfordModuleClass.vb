Option Explicit

Private benfordCheckValues As Variant
Private benfordCount As Long

Sub Class_Initialize()

    Dim counter As Long
    ReDim benfordCheckValues(9)

    For counter = LBound(benfordCheckValues) To UBound(benfordCheckValues)
        benfordCheckValues(counter) = 0
    Next counter

End Sub

Function InitialValuesBenford(val As Long) As Double
        
    '1 = "30,1%"
    '2 = "17,6%"
    '3 = "12,5%"
    '4 = " 9,7%"
    '5 = " 7,9%"
    '6 = " 6,7%"
    '7 = " 5,8%"
    '8 = " 5,1%"
    '9 = " 4,6%"
    
    InitialValuesBenford = Round(WorksheetFunction.Log10(1 + 1 / val), 3)
    
End Function

Function PercentageFixer(valToReturn As Double) As String
                    
    If valToReturn > 0.1 Then
        PercentageFixer = Trim(Format(valToReturn, "##.0%"))
    ElseIf valToReturn = 0 Then
        PercentageFixer = " " & Format(valToReturn, "0.0%")
    Else
        PercentageFixer = " " & Format(valToReturn, "#.0%")
    End If
    
End Function

Function CreateBenfordLawReport() As String

    Dim line As String: line = "---------------------------------"
    On Error GoTo CreateBenfordLawReport_Error

    Dim counter      As Long
    CreateBenfordLawReport = line & line & line & vbCrLf _
                            & line & line & line & vbCrLf _
                            & line & line & line & vbCrLf _
                            & "Benford's Law" & vbCrLf & "https://en.wikipedia.org/wiki/Benford%27s_law" & vbCrLf

    For counter = LBound(CheckValues) To UBound(CheckValues)
        If counter = 0 Then
            Dim header As String
            header = CreateBenfordLawReport & vbCrLf & "#" & vbTab & _
                                    "-> " & "Val." & vbTab & "Real%" & vbTab & "Expected"
            CreateBenfordLawReport = header
        Else
            CreateBenfordLawReport = CreateBenfordLawReport & vbCrLf & counter & vbTab & _
                                    "-> " & CheckValues(counter) & vbTab & _
                                    PercentageFixer(Round(CheckValues(counter) / Me.Count, 3)) & vbTab & _
                                    PercentageFixer(InitialValuesBenford(counter)) & vbTab & "|"
        End If
        
        If counter = 0 Or counter = 9 Then
            CreateBenfordLawReport = CreateBenfordLawReport & vbCrLf & line
        End If
    Next counter

    On Error GoTo 0
    Exit Function

CreateBenfordLawReport_Error:

    CreateBenfordLawReport = "Not enough data..."

End Function

Property Get CheckValues() As Variant
    CheckValues = benfordCheckValues
End Property

Property Get Count() As Long
    Count = benfordCount
End Property

Sub IncrementCount()
    benfordCount = benfordCount + 1
End Sub

Sub IncrementValue(valToInput As Variant)

    Dim leftDigit As Variant
    leftDigit = Left(valToInput, 1)
    benfordCheckValues(leftDigit) = benfordCheckValues(leftDigit) + 1
    
End Sub
