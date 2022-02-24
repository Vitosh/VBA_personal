Option Explicit

Public Sub CreateLogFile(Optional report As String)

    On Error GoTo CreateLogFile_Error
    
    WaitASecond
    Dim newFilePath As String
    newFilePath = "\reports"
    Dim fileName As String
    fileName = ThisWorkbook.Path & newFilePath & CodifyMyTime(True)
    If Dir(ThisWorkbook.Path & newFilePath, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFilePath
    
    Dim fs  As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Dim notepad As Object
    Set notepad = fs.CreateTextFile(fileName, True)

    notepad.WriteLine report
    notepad.Close
    
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:
    Debug.Print "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"
End Sub

Public Function CodifyMyTime(Optional makepath As Boolean = False) As String

    On Error GoTo codify_Error

    Dim timePart01 As Double
    Dim timePart02 As Double
    Dim timePartNow As Double

    timePartNow = Round(Now(), 8)
    timePart01 = Split(CStr(timePartNow), ".")(0)
    timePart02 = Split(CStr(timePartNow), ".")(1)
    CodifyMyTime = Format(Now, "YYYYMMMDD_HHNNSS") & "_" & Hex(timePart01) & "_" & Hex(timePart02)

    If makepath Then CodifyMyTime = "\" & CodifyMyTime & ".xml"
    On Error GoTo 0
    Exit Function

codify_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CodifyTime"
End Function                        
                        
Public Sub WaitASecond()
    Application.Wait (Now + TimeValue("00:00:01"))
End Sub
