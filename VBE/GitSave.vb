Option Explicit

Sub GitSave()

    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.codeModule.lines(1, item.codeModule.CountOfLines)
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.path & "\VBA\VBA-Code_Together\"
    Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.path & "\VBA\VBA-Code_By_Modules\"

    Kill pathToExport & "*.*"
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
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
            component.export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Function GetFolderOnDesktopPath() As String

    Dim shell As Object
    Dim fso As Object
    Dim specialFolderPath As String

    Set shell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    specialFolderPath = shell.SpecialFolders("Desktop")
    If Right(specialFolderPath, 1) <> "\" Then specialFolderPath = specialFolderPath & "\"
    
    GetFolderOnDesktopPath = specialFolderPath & Split(ThisWorkbook.Name, "_")(0) & "\"
    
End Function

Sub CreateFolderOnDesktop(specialFolderPath As String)
    
    On Error Resume Next
    
    MkDir specialFolderPath
    If Err.Number <> 0 Then
        If Err.Number = 75 Then
            Debug.Print "Folder exists - " & specialFolderPath
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    Else
        Debug.Print "Folder has been created - " & specialFolderPath
    End If
    
    On Error GoTo 0
    
End Sub
