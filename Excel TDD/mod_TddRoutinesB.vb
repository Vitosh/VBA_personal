Option Explicit
Option Private Module

Private lngCol              As Long
Private lngRow              As Long
Private lngCounter          As Long

Public Sub Tdd_01()

    On Error Resume Next

    Dim specs               As New SpecSuite

    Dim lngValue            As Long
    Dim dtValue             As Date
    Dim strInitial          As String

    Call OnStart
    
    specs.It("001", "Just A Test").Expect(2).ToEqual 1 + 1
    specs.It("002", "Just A Test").Expect(2).ToNotEqual 1 + 1 + 2
    
    InlineRunner.RunSuite specs
    Call specs.TotalTests

    Call OnEnd
    
    On Error GoTo 0

End Sub

