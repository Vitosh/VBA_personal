Attribute VB_Name = "Modul1"
Option Explicit

Public Sub Main()

    Dim totalTests As Long
    Dim pathInputTests As String
    Dim pathOutputTests As String

    Dim inputTests As Variant
    Dim outputTests As Variant

    Dim cntTests As Long
    Dim cnt As Long

    pathInputTests = "C:\Desktop\Test002.txt"
    pathOutputTests = "C:\Desktop\Result002.txt"

    inputTests = Split(ReadFileLineByLineToString(pathInputTests), vbCrLf)
    outputTests = Split(ReadFileLineByLineToString(pathOutputTests), vbCrLf)

    For cnt = LBound(inputTests) To UBound(inputTests)

        Dim expectedValue   As Variant
        Dim receivedValue   As Variant

        On Error Resume Next

        expectedValue = outputTests(cnt)
        receivedValue = MainTest(Trim(inputTests(cnt)))

        If Err.Number <> 0 Then
            Debug.Print runtimeError(cnt)
            Err.Clear
        Else
            If Trim(expectedValue) = Trim(receivedValue) Then
                Debug.Print positiveResult(cnt)
            Else
                Debug.Print negativeResult(cnt, expectedValue, receivedValue)
            End If
        End If

    Next cnt

End Sub

Public Function runtimeError(ByVal cnt As Long) As String
    cnt = cnt + 1
    runtimeError = "Runtime error on " & cnt & "!"
End Function

Public Function positiveResult(ByVal cnt As Long) As String
    cnt = cnt + 1
    positiveResult = "Test " & cnt & "..................................... ok!"
End Function

Public Function negativeResult(ByVal cnt As Long, expected As Variant, _
                                                received As Variant) As String
    cnt = cnt + 1
    negativeResult = "Error on test " & cnt & "!" & _
                    " Expected -> " & vbTab & expected & vbTab & _
                    " Received -> " & vbTab & received

End Function

'---------------------------------------------------------------------------------------
' Method : MainTest
' Purpose: This is where the competitors paste their solution.
'---------------------------------------------------------------------------------------

Public Function MainTest(ByVal consoleInput As String) As String

    Dim inputVar    As Variant
    Dim cnt         As Long
    Dim outputVar   As Variant
        
    inputVar = Split(consoleInput)
    ReDim outputVar(UBound(inputVar))
    
    For cnt = LBound(inputVar) To UBound(inputVar)
        If Asc(inputVar(cnt)) = Asc("z") Then
            MainTest = MainTest & " a"
        Else
            MainTest = MainTest & " " & Chr(Asc(inputVar(cnt)) + 1)
        End If
        
    Next cnt

'    Dim a   As Double
'    Dim b   As Double
'    Dim c   As Double
'
'    a = Split(consoleInput)(0)
'    b = Split(consoleInput)(1)
'    c = Split(consoleInput)(2)
'
'    If c Mod 2 = 0 Then
'        MainTest = a + b + c
'    Else
'        MainTest = a + b - c
'    End If

End Function


Public Function ReadFromFile(path As String) As String

    Dim fileNo As Long
    fileNo = FreeFile

    Open path For Input As #fileNo

    Do While Not EOF(fileNo)
        Dim textRowInput As String
        Line Input #fileNo, textRowInput
        ReadFromFile = ReadFromFile & textRowInput
        If Not EOF(fileNo) Then
            ReadFromFile = ReadFromFile & vbCrLf
        End If
    Loop

    Close #fileNo

End Function

Sub WriteToFile(filePath As String, text As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(filePath)
    oFile.Write text
    oFile.Close
    
End Sub

Sub TestMe()

    Dim readTxt As String
    Dim filePath As String: filePath = "C:\text.txt"

    readTxt = ReadFromFile(filePath)
    readTxt = Replace(readTxt, "name=", "")
    readTxt = Replace(readTxt, "correo=", "")

    WriteToFile filePath, readTxt

End Sub



