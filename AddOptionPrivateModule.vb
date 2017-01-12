Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Method : AddOptionPrivate
' Author : stackoverflow.com
' Date   : 12.01.2017
' Purpose: Checking for "Option Private Mod~" up to line 5, if not found we add it in
'           every module
'---------------------------------------------------------------------------------------
Sub AddOptionPrivate()

    Const UP_TO_LINE = 5
    Const PRIVATE_MODULE = "Option Private Module"

    Dim objXL               As Object
    
    Dim objPro              As Object
    Dim objComp             As Variant
    Dim strText             As String
     
    Set objXL = GetObject(, "Excel.Application")
    Set objPro = objXL.ActiveWorkbook.VBProject
    
    For Each objComp In objPro.VBComponents
        If objComp.Type = 1 Then
            strText = objComp.CodeModule.Lines(1, UP_TO_LINE)
            
            If InStr(1, strText, PRIVATE_MODULE) = 0 Then
                objComp.CodeModule.InsertLines 2, PRIVATE_MODULE
            End If
            
        End If
    Next objComp
    
End Sub
