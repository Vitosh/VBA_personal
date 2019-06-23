Option Explicit

Public Sub MainBenfordCheck(myRange As Range)
    
    Dim myCell     As Range
    Dim benford    As New BenfordModel
            
    For Each myCell In myRange
        If IsNumeric(myCell) Then
            benford.IncrementValue Abs(myCell.value)
            benford.IncrementCount
        End If
    Next myCell
    
    CreateLogFile benford.CreateBenfordLawReport
    
End Sub

Public Sub CreateLogFile(Optional report As String)

    On Error GoTo CreateLogFile_Error
    
    Dim newFilePath As String
    newFilePath = "\tests_info"
     
    Dim filename As String
    filename = ThisWorkbook.Path & newFilePath & CodifyTime(True)
    If Dir(ThisWorkbook.Path & newFilePath, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFilePath
    
    Dim fs  As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim notepad As Object
    Set notepad = fs.CreateTextFile(filename, True)

    Dim header  As String
    header = Now & vbCrLf & "Created by: " & Environ("USERNAME")
    
    notepad.WriteLine header
    notepad.WriteLine report
    notepad.Close
    
    Dim shellCommand        As String
    shellCommand = "C:\WINDOWS\notepad.exe "
    shellCommand = shellCommand & filename
    Shell shellCommand

    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub

Public Function CodifyTime(Optional makePath As Boolean = False) As String

    On Error GoTo codify_Error

    Dim timePart01 As Double
    Dim timePart02 As Double
    Dim timePartNow As Double

    timePartNow = Round(Now(), 8)
    timePart01 = Split(CStr(timePartNow), ",")(0)
    timePart02 = Split(CStr(timePartNow), ",")(1)
    CodifyTime = Hex(timePart01) & "_" & Hex(timePart02)

    If makePath Then CodifyTime = "\" & CodifyTime & ".txt"

    On Error GoTo 0
    Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function

