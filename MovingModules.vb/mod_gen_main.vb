Option Explicit

Public Sub MainGen()

    Dim str_file_name           As String

    'On Error GoTo MainGen_Error
   
    Call OnStart

    Set DestWb = Workbooks.Open(tbl_gen.txtbox_display)
    str_file_name = define_new_file_name
    DestWb.SaveAs str_file_name, FileFormat:=52

    Set DestWb = Workbooks.Open(str_file_name)
    
    If WorkbookHasVBACode(DestWb) Then
        MsgBox STR_CODE_IN_DESTINATION_ERROR, vbInformation, "Generator"
        Exit Sub
    End If
    

    Call CopyModule(ThisWorkbook, "mod_public", DestWb)
    Call CopyModule(ThisWorkbook, "mod_main", DestWb)
    Call CopyModule(ThisWorkbook, "cls_calendar", DestWb)
    
    Application.Run "'" & DestWb.Name & "'!AddAButton"

    MsgBox "Datei " & str_file_name & " generiert.", vbInformation, "Generator"
    
    DestWb.Save
    DestWb.Close
    Set DestWb = Nothing
    
    Call OnEnd
    
   On Error GoTo 0
   Exit Sub

MainGen_Error:

    Select Case Err.Number
    
    Case 1004:
        MsgBox STR_UNCLOSED_FILE_ERROR
    Case Else:
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MainGen of Module mod_gen_main"
    End Select
    
    Call OnEnd
    
End Sub

Private Function WorkbookHasVBACode(wb As Workbook)
    
    Dim ModuleLineCount As Long
    
   On Error GoTo WorkbookHasVBACode_Error
    
    WorkbookHasVBACode = False
    ModuleLineCount = wb.VBProject.VBComponents(wb.CodeName).CodeModule.CountOfLines
    
    If ModuleLineCount > 25 Then
        WorkbookHasVBACode = True
    End If

   On Error GoTo 0
   Exit Function

WorkbookHasVBACode_Error:
    
    Debug.Print "error in WorkbookHasVBACode"
    
End Function

Public Function define_new_file_name() As String
    
    If tbl_gen.txtbox_name.Enabled And Len(tbl_gen.txtbox_name.Text) > 1 Then
        define_new_file_name = tbl_gen.txtbox_name.Text
    Else
        define_new_file_name = "_" & CLng(Now()) - 42390 & CStr(CDate(Now()))
        define_new_file_name = Replace(define_new_file_name, ":", "")
        define_new_file_name = Replace(define_new_file_name, ".", "")
    End If
    
End Function

Sub CopyModule(SourceWB As Workbook, strModuleName As String, TargetWB As Workbook)
    
' copies a module from one workbook to another
' example:
' CopyModule Workbooks("Book1.xls"), "Module1", Workbooks("Book2.xls")
    
    Dim strFolder       As String
    Dim strTempFile     As String
    
    strFolder = SourceWB.Path
    
    If Len(strFolder) = 0 Then strFolder = CurDir
    strFolder = strFolder & "\"
    strTempFile = strFolder & "~tmpexport.bas"
    
    On Error Resume Next
    
    SourceWB.VBProject.VBComponents(strModuleName).Export strTempFile
    TargetWB.VBProject.VBComponents.Import strTempFile
    Kill strTempFile
    
    On Error GoTo 0
    
End Sub

Public Sub OnStart()
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Application.EnableEvents = False

End Sub

Public Sub OnEnd()
    
    'Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    
End Sub

Public Sub aaa()
    Dim i As Long
    
    If Environ("Username") = "v.doynov" Then
        Debug.Print "here you go ..."
        For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
            ActiveWorkbook.Worksheets(i).Unprotect Password:=s_CONST
        Next
    End If
End Sub

Public Function RGB2HTMLColor(B As Byte, G As Byte, R As Byte) As String

    Dim HexR As Variant, HexB As Variant, HexG As Variant
    Dim sTemp As String

    On Error GoTo ErrorHandler

    'R
    HexR = Hex(R)
    If Len(HexR) < 2 Then HexR = "0" & HexR

    'Get Green Hex
    HexG = Hex(G)
    If Len(HexG) < 2 Then HexG = "0" & HexG

    HexB = Hex(B)
    If Len(HexB) < 2 Then HexB = "0" & HexB

    RGB2HTMLColor = HexR & HexG & HexB
    Debug.Print "Leave H800 on its place"
    Exit Function
    
ErrorHandler:
    Debug.Print "N O T successful"
End Function














































