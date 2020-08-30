Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.codeModule.Lines(1, item.codeModule.CountOfLines)
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
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
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
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
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub
