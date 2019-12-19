Attribute VB_Name = "ExcelPrintToNotepad"
Option Explicit
Option Private Module

Sub PrintToNotepad(Optional dataToPrint As String = "")

    If SET_IN_PRODUCTION Then On Error GoTo CreateLogFile_Error
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim filename As String
    Dim newFile  As String
    Dim str_shell  As String
    
    newFile = "\Info"
    
    filename = ThisWorkbook.path & newFile & CodifyTime(True)
    If Dir(ThisWorkbook.path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(filename, True)
    
    If dataToPrint <> "" Then
        textObject.WriteLine dataToPrint
    Else
        textObject.WriteLine PUB_STR_ERROR_REPORT
    End If
    textObject.Close
    
    str_shell = "C:\WINDOWS\notepad.exe "
    str_shell = str_shell & filename
    shell str_shell
    
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

Public Sub MakeAllValues()
    
    Dim myCell                 As Range
    Dim i               As Long
    Dim str                     As String
    
    For Each myCell In Selection
        Increment i
        str = vbTab & "my_arr(" & i & ")= "
        
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
