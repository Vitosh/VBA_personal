Option Explicit

Private m_CheckValues       As Variant
Private m_lCount            As Long

Private Sub Class_Initialize()

    Dim lngCounter As Long
    ReDim m_CheckValues(10)

    For lngCounter = LBound(m_CheckValues) To UBound(m_CheckValues)
        m_CheckValues(lngCounter) = 0
    Next lngCounter

End Sub

Public Function ValuesBenford() As String
    
    Dim cnt As Long
    ValuesBenford = ""
    
    For cnt = 1 To 9
        ValuesBenford = ValuesBenford & Round(WorksheetFunction.Log10(1 + 1 / cnt), 4) & vbCrLf
    Next cnt
    
End Function

Public Function CalculateBenfordValue(lngDouble As Long) As Double
    
    CalculateBenfordValue = Round(WorksheetFunction.Log10(1 + 1 / lngDouble), 3)
    
End Function

Public Function InitialValuesBenford(lngVal As Long) As Variant

    ReDim varB(10)
    varB(0) = "  N/A"
    varB(1) = "30,1%"
    varB(2) = "17,6%"
    varB(3) = "12,5%"
    varB(4) = " 9,7%"
    varB(5) = " 7,9%"
    varB(6) = " 6,7%"
    varB(7) = " 5,8%"
    varB(8) = " 5,1%"
    varB(9) = " 4,6%"
    varB(10) = "  N/A"
    InitialValuesBenford = varB(lngVal)

End Function

Public Function CreateBenfordLawReport() As String

    Dim strLine As String: strLine = "---------------------------------"
    On Error GoTo CreateBenfordLawReport_Error

    Dim lngCounter      As Long
    CreateBenfordLawReport = vbCrLf _
                                & strLine & strLine & strLine & vbCrLf _
                                & strLine & strLine & strLine & vbCrLf _
                                & strLine & strLine & strLine & vbCrLf _
                                & "Benford's Law" & vbCrLf & "https://en.wikipedia.org/wiki/Benford%27s_law" & vbCrLf

    For lngCounter = LBound(CheckValues) To UBound(CheckValues)
        CreateBenfordLawReport = CreateBenfordLawReport & vbCrLf _
                                & lngCounter & vbTab & _
                                "-> " & CheckValues(lngCounter) & vbTab & _
                                Round(CheckValues(lngCounter) / Me.Count, 3) * 100 & "%" & vbTab & _
                                Me.InitialValuesBenford(lngCounter) & vbTab & "|"
                
        If lngCounter = 0 Or lngCounter = 9 Then
            CreateBenfordLawReport = CreateBenfordLawReport & vbCrLf & strLine
        End If
        
    Next lngCounter

    CreateBenfordLawReport = CreateBenfordLawReport & vbCrLf & IIf(Len(STR_UNKNOWN_VALUES), vbCrLf & "Unknown values:" & vbCrLf & STR_UNKNOWN_VALUES, "")
    CreateBenfordLawReport = CreateBenfordLawReport & vbCrLf & IIf(Len(STR_ZERO_VALUES), "Zero values:" & vbCrLf & STR_ZERO_VALUES, "")

    On Error GoTo 0
    Exit Function

CreateBenfordLawReport_Error:

    CreateBenfordLawReport = "Not enough data..."

End Function

Public Property Get CheckValues() As Variant

    CheckValues = m_CheckValues

End Property

Public Property Get Count() As Long

    Count = m_lCount

End Property

Public Sub IncrementCount()

    m_lCount = m_lCount + 1

End Sub

Public Sub IncrementValue(varInput As Variant, strWksName As String, strInvoiceNumber As String)

    Dim varInputLeft    As Variant

    varInputLeft = Left(varInput, 1)

    If IsNumeric(varInputLeft) Then
        m_CheckValues(varInputLeft) = m_CheckValues(varInputLeft) + 1
        If varInputLeft > 0 Then
            Me.IncrementCount
        Else
            STR_ZERO_VALUES = STR_ZERO_VALUES & varInput & vbTab & vbTab & strWksName & vbTab & vbTab & strInvoiceNumber & vbCrLf
        End If
    Else
        m_CheckValues(10) = m_CheckValues(10) + 1
        STR_UNKNOWN_VALUES = STR_UNKNOWN_VALUES & varInput & vbTab & vbTab & strWksName & vbTab & vbTab & strInvoiceNumber & vbCrLf
    End If

End Sub




