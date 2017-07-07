Option Explicit
Private pSpecsCol As Collection

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Public Description              As String
Public BeforeEachCallback       As String
Public BeforeEachCallbackArgs   As Variant
Private pCounter                As Long

Public Property Get SpecsCol() As Collection

    If pSpecsCol Is Nothing Then: Set pSpecsCol = New Collection
    Set SpecsCol = pSpecsCol
    
End Property
Public Property Let SpecsCol(value As Collection)
    
    Set pSpecsCol = value
    
End Property


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Public Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

''
' Create a new spec definition with description
'
' @param {String} Description
' @param {String} [SpecId] Useful for identifying specific specs
' @returns {SpecDefinition} Initialized Spec Definition
' --------------------------------------------- '

Public Function It(Description As String, Optional SpecId As String = "") As SpecDefinition
    
    Dim Spec As New SpecDefinition
    
    pCounter = pCounter + 1
    ' Call BeforeEach if defined
    ExecuteBeforeEach
    
    ' Initialize spec
    Spec.Description = Description
    Spec.ID = SpecId
    Me.SpecsCol.Add Spec
    Set It = Spec
    
End Function

Public Function f_lng_number_tests() As Long
    f_lng_number_tests = pCounter
End Function

Public Sub TotalTests()
    
    Call Increment(LNG_TOTAL_TESTS, Me.f_lng_number_tests)
    Debug.Print "  Tests:" & pCounter & vbCrLf
    STR_ERROR_REPORT = STR_ERROR_REPORT & vbCrLf & "  Tests:" & pCounter & vbCrLf & vbCrLf
 
End Sub

Public Sub BeforeEach(Callback As String, ParamArray CallbackArgs() As Variant)
    Me.BeforeEachCallback = Callback
    Me.BeforeEachCallbackArgs = CallbackArgs
End Sub


' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' Internal Methods
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Private Sub ExecuteBeforeEach()

    If Me.BeforeEachCallback <> "" Then
        Dim HasArguments As Boolean
        If VarType(Me.BeforeEachCallbackArgs) = vbObject Then
            If Not Me.BeforeEachCallbackArgs Is Nothing Then
                HasArguments = True
            End If
        ElseIf IsArray(Me.BeforeEachCallbackArgs) Then
            If UBound(Me.BeforeEachCallbackArgs) >= 0 Then
                HasArguments = True
            End If
        End If
    
        If HasArguments Then
            Application.Run Me.BeforeEachCallback, Me.BeforeEachCallbackArgs
        Else
            Application.Run Me.BeforeEachCallback
        End If
    End If
    
End Sub

