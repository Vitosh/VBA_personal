Attribute VB_Name = "ExcelVBE"
Option Explicit
Option Private Module

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.vbProject.VBComponents
        lineToPrint = item.codeModule.lines(1, item.codeModule.CountOfLines)
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    PrintToNotepad textToPrint
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.vbProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    PrintToNotepad textToPrint
    
End Sub

Sub ListProcedures(Optional modName As String = "ExcelAdditional", Optional withParentInfo As Boolean = False)
    
    Dim project As VBIDE.vbProject
    Dim component As VBIDE.VBComponent
    Dim codeModule As VBIDE.codeModule
    Dim lineNum As Long
    Dim procName As String
    Dim procKind As VBIDE.vbext_ProcKind
    Dim subsInfo As String
    
    Set project = ThisWorkbook.vbProject
    Set component = project.VBComponents(modName)
    Set codeModule = component.codeModule

    With codeModule
        lineNum = .CountOfDeclarationLines + 1
        
        Do Until lineNum >= .CountOfLines
            procName = .ProcOfLine(lineNum, procKind)

            If withParentInfo Then
                subsInfo = subsInfo & IIf(subsInfo = vbNullString, vbNullString, vbCrLf) & modName & "." & procName
            Else
                subsInfo = subsInfo & IIf(subsInfo = vbNullString, vbNullString, vbCrLf) & procName
            End If

            lineNum = .ProcStartLine(procName, procKind) + .ProcCountLines(procName, procKind) + 1
        Loop
        
    End With
    
    Debug.Print subsInfo
    PrintToNotepad subsInfo
    
End Sub

Sub ExportModules()
    
    CreateFolderOnDesktop GetFolderOnDesktopPath
    
    On Error Resume Next
    Kill GetFolderOnDesktopPath & "\*.*"
    On Error GoTo 0
    
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    If wkb.vbProject.Protection = vbext_pp_locked Then
        Debug.Print "The VBA in this workbook is locked."
        Exit Sub
    End If
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.vbProject.VBComponents
        tryExport = True
        filePath = component.Name
        
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Increment unitsCount
            Debug.Print unitsCount & " exporting " & filePath
            component.export GetFolderOnDesktopPath & filePath
        End If
        
    Next

    Debug.Print "Exported at " & GetFolderOnDesktopPath
    
End Sub

Function GetFolderOnDesktopPath() As String

    Dim shell As Object
    Dim fso As Object
    Dim specialFolderPath As String

    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    specialFolderPath = shell.SpecialFolders("Desktop")
    If Right(specialFolderPath, 1) <> "\" Then specialFolderPath = specialFolderPath & "\"
    
    GetFolderOnDesktopPath = specialFolderPath & CON_STR_APP_NAME & "\"
    
End Function

Sub CreateFolderOnDesktop(specialFolderPath As String)
    
    On Error Resume Next
    
    MkDir specialFolderPath
    If Err.Number <> 0 Then
        If Err.Number = 75 Then
            Debug.Print "Folder exists - " & specialFolderPath
        Else
            Err.Raise Err.Number, Err.source, Err.Description
        End If
    Else
        Debug.Print "Folder has been created - " & specialFolderPath
    End If
    
    On Error GoTo 0
    
End Sub

Public Sub ImportModules()
    
    '1. The target workbook should be opened in the same Excel instance as the ThisWorkbook
    '2. The target workbook should be in the same directory as ThisWorkbook
    '3. The code to be added should be present in GetFolderOnDesktopPath
    
    Dim targetName As String: targetName = "empty.xlsm"
    Dim targetPath As String: targetPath = ThisWorkbook.path & "\" & targetName
    
    Dim wkbTarget As Workbook
    Dim fso As Scripting.FileSystemObject
    Dim file As Scripting.file
    Dim codePath As String: codePath = GetFolderOnDesktopPath
  
    Set wkbTarget = Workbooks(targetName)
    
    If wkbTarget.vbProject.Protection = 1 Then
        Debug.Print "VBProject is protected!"
    End If
    
    Set fso = New Scripting.FileSystemObject
    If fso.GetFolder(codePath).Files.Count = 0 Then
       Debug.Print "Zero vba files in source workbook!"
       Exit Sub
    End If
    
    DeleteAllVba wkbTarget

    Dim unitsCount As Long
    For Each file In fso.GetFolder(codePath).Files
        Select Case fso.GetExtensionName(file.Name)
            Case "cls", "frm", "bas":
                Increment unitsCount
                Debug.Print unitsCount & " -> in " & wkbTarget.Name & " adding " & file.Name
                wkbTarget.vbProject.VBComponents.Import file.path
            Case Else:
                Debug.Print file.Name & " cannot be processed."
        End Select
    Next
    
    Debug.Print vbCrLf & unitsCount & " units were just added to:" & vbCrLf & targetPath
    
End Sub

Function DeleteAllVba(wkbTarget As Workbook)

        Dim project As VBIDE.vbProject
        Dim component As VBIDE.VBComponent
        Dim unitsCount As Long
        
        Set project = wkbTarget.vbProject
        
        For Each component In project.VBComponents
            If component.Type <> vbext_ct_Document Then
                Increment unitsCount
                Debug.Print unitsCount & " from " & wkbTarget.Name & " deleting " & component.Name
                project.VBComponents.Remove component
            End If
        Next
         
        Debug.Print 'Empty line is good :)
        
End Function

