VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tddSpecSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pSpecsCol As Collection

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

Public Function It(Description As String, Optional SpecId As String = "") As tddSpecDefinition
    
    Dim Spec As New tddSpecDefinition
    
    pCounter = pCounter + 1
    ExecuteBeforeEach
    Spec.Description = Description
    Spec.Id = SpecId
    Me.SpecsCol.Add Spec
    Set It = Spec
    
End Function

Public Sub TotalTests()

    Debug.Print "Total tests:" & pCounter
    PUB_STR_ERROR_REPORT = PUB_STR_ERROR_REPORT & "Total tests:" & pCounter & vbCrLf
 
End Sub

Public Sub BeforeEach(Callback As String, ParamArray CallbackArgs() As Variant)
    Me.BeforeEachCallback = Callback
    Me.BeforeEachCallbackArgs = CallbackArgs
End Sub

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

