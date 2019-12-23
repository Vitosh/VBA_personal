Attribute VB_Name = "ExcelPrintToNotepad"
Option Explicit
Option Private Module

Sub PrintToNotepad(Optional dataToPrint As String = "")

    If SET_IN_PRODUCTION Then On Error GoTo CreateLogFile_Error
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String

    newFile = "\Info"
    
    fileName = ThisWorkbook.path & newFile & CodifyTime(True)
    If Dir(ThisWorkbook.path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(fileName, True)
    
    If dataToPrint <> "" Then
        textObject.WriteLine dataToPrint
    Else
        textObject.WriteLine PUB_STR_ERROR_REPORT
    End If
    
    textObject.Close
    
    shellPath = "C:\WINDOWS\notepad.exe "
    shellPath = shellPath & fileName
    shell shellPath
    
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub

Public Function CodifyTime(Optional makeString As Boolean = False) As String

    If SET_IN_PRODUCTION Then On Error GoTo codify_Error
    
    Dim leftPart                  As Variant
    Dim rightPart                  As Variant
    Dim initialTime                 As Double
    
    initialTime = Round(Now(), 8)
    
    leftPart = Split(CStr(initialTime), ".")(0)
    rightPart = Split(CStr(initialTime), ".")(1)
    
    CodifyTime = Hex(leftPart) & "_" & Hex(rightPart)
    
    If makeString Then CodifyTime = "\" & CodifyTime & ".txt"
    
    On Error GoTo 0
    Exit Function

codify_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure codify of Function TDD_Export"

End Function

Public Function DecodifyTime(hexTime As String) As String
    
    Dim leftPart                  As Variant
    Dim rightPart                  As Variant
    
    leftPart = Split(hexTime, "_")(0)
    rightPart = Split(hexTime, "_")(1)
    
    DecodifyTime = CLng("&H" & leftPart) & "." & CLng("&H" & rightPart)
    
End Function
