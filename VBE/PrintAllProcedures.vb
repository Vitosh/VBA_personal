'---------------------------------------------------------------------------------------
' Purpose   :       Prints all subs and functions in a project
' Prerequisites:    Microsoft Visual Basic for Applications Extensibility 5.3 library
'                   CreateLogFile
' How to run:       Run GetFunctionAndSubNames, set a parameter to blnWithParentInfo
'---------------------------------------------------------------------------------------

Option Explicit

Private strSubsInfo As String

Public Sub GetFunctionAndSubNames()
    
    Dim item            As Variant
    
    strSubsInfo = ""
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        ListProcedures item.name, False
        'Debug.Print item.CodeModule.lines(1, item.CodeModule.CountOfLines)
    Next item
    
    CreateLogFile strSubsInfo
    
End Sub

Private Sub ListProcedures(strName As String, Optional blnWithParentInfo = False)

    'Microsoft Visual Basic for Applications Extensibility 5.3 library

    Dim VBProj          As VBIDE.VBProject
    Dim VBComp          As VBIDE.VBComponent
    Dim CodeMod         As VBIDE.CodeModule
    Dim LineNum         As Long
    Dim ProcName        As String
    Dim ProcKind        As VBIDE.vbext_ProcKind

    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents(strName)
    Set CodeMod = VBComp.CodeModule

    With CodeMod
        LineNum = .CountOfDeclarationLines + 1
        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, ProcKind)

            If blnWithParentInfo Then
                strSubsInfo = strSubsInfo & IIf(strSubsInfo = vbNullString, vbNullString, vbCrLf) & strName & "." & ProcName
            Else
                strSubsInfo = strSubsInfo & IIf(strSubsInfo = vbNullString, vbNullString, vbCrLf) & ProcName
            End If

            LineNum = .ProcStartLine(ProcName, ProcKind) + .ProcCountLines(ProcName, ProcKind) + 1
        Loop
    End With

End Sub
