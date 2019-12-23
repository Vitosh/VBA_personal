Attribute VB_Name = "ExcelStructure"
Option Explicit
Option Private Module

Public Sub LockScroll(lockArea As Range)
    
    Dim wks As Worksheet
    For Each wks In ThisWorkbook.Worksheets
        wks.ScrollArea = lockArea.Address
    Next wks
    
End Sub

Public Sub UnlockScroll()
    
    Dim wks As Worksheet
    For Each wks In ThisWorkbook.Worksheets
        wks.ScrollArea = ""
    Next wks
    
End Sub

Sub StyleKiller()

    Dim myStyle As Style
    
    For Each myStyle In ThisWorkbook.Styles
        If Not myStyle.BuiltIn Then
            Debug.Print myStyle.Name
            myStyle.Delete
        End If
    Next

End Sub

Public Sub DeleteName(myName As String)

    On Error GoTo DeleteName_Error

    ThisWorkbook.Names(myName).Delete
    Debug.Print myName & " is deleted!"
    
    On Error GoTo 0
    Exit Sub

DeleteName_Error:

    Debug.Print myName & " not present or some error"
    On Error GoTo 0
    
End Sub

Sub CoverRange(myRange As Range, wks As Worksheet)
    
    Dim myLeft As Long
    Dim myTop As Long
    Dim myWidth As Long
    Dim myHeight As Long
    
    If wks.Name <> ActiveSheet.Name Then
        MsgBox "You better select the sheet you are working on..."
        Exit Sub
    End If
    
    myLeft = myRange.Left
    myTop = myRange.Top
    myWidth = myRange.Width
    myHeight = myRange.Height
    
    With wks.Shapes
        .AddTextbox(msoTextOrientationVertical, myLeft, myTop, myWidth, myHeight).Select
        Selection.ShapeRange.Line.Visible = msoFalse
    End With

End Sub

Public Sub PrintSheetPDF(inputPrintArea As Range, _
                                printedFileName As String, _
                                Optional isBlack As Boolean = False)

    If SET_IN_PRODUCTION Then On Error GoTo PrintPDF_Error
    
    Dim wks As Worksheet
    Set wks = Worksheets(inputPrintArea.Parent.Name)
    
    With wks
        .PageSetup.Zoom = False
        .PageSetup.BlackAndWhite = isBlack

        inputPrintArea.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=printedFileName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    End With

    On Error GoTo 0
    Exit Sub

PrintPDF_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPDF of Modul mod_Drucken"

End Sub

Public Sub PrintPage(printRange As Range, Optional isBlack As Boolean = False)

    Dim wksSheet As Worksheet
    Dim reducePaperTitle As String

    On Error GoTo PrintPage_Error

    reducePaperTitle = "Reduce printing and save trees!"
    printRange.Parent.PageSetup.BlackAndWhite = isBlack

    Set wksSheet = printRange.Parent

    With wksSheet.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With

    Select Case MsgBox("Are you sure you would like to print the selected page?", vbYesNo Or vbQuestion Or vbDefaultButton1, reducePaperTitle)
        Case vbYes
            Select Case MsgBox("Really?", vbYesNo Or vbQuestion Or vbDefaultButton1, reducePaperTitle)
                Case vbYes
                    printRange.PrintOut
            End Select
    End Select

    On Error GoTo 0
    Exit Sub

PrintPage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PrintPage of Modul mod_Drucken"

End Sub

Sub DeleteDrawingObjects(wks As Worksheet)

    Dim i           As Long
    
    For i = wks.DrawingObjects().Count To 1 Step -1
        wks.DrawingObjects(i).Delete
    Next i

End Sub

Public Sub UnhideAll()

    Dim wks As Worksheet

    For Each wks In ThisWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next

    UnprotectAll

End Sub

Public Sub UnprotectAll()

    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        ThisWorkbook.Worksheets(i).Unprotect Password:=WORKSHEET_UNPROTECT_PASSWORD
    Next i
    
End Sub

Public Sub HideNeededWorksheets()

    Dim varSheet As Variant
    Dim visibleSheets As Variant
    Dim hiddenSheets As Variant

    OnStart

    visibleSheets = Array(tblInput)
    hiddenSheets = Array(tblSettings)

    For Each varSheet In visibleSheets
        varSheet.Visible = xlSheetVisible
    Next varSheet

    For Each varSheet In hiddenSheets
        varSheet.Visible = xlSheetVeryHidden
    Next varSheet

    OnEnd

End Sub

Public Sub AddCommentToSelection(myComment As Range)
    
    Dim myCell            As Range

    For Each myCell In Selection
             myCell.ClearComments
            myCell.AddComment myComment.Text
            myCell.Comment.Visible = False
            myCell.Comment.Shape.ScaleWidth 4, msoFalse, msoScaleFromTopLeft
            myCell.Comment.Shape.ScaleHeight 1.5, msoFalse, msoScaleFromTopLeft

    Next myCell

End Sub

Public Sub PrintArray(myArray As Variant)

    Dim i As Long
    For i = LBound(myArray) To UBound(myArray)
        Debug.Print i & " --> " & myArray(i)
    Next i
    
End Sub

Sub PrintAllNames()
    
    Dim nm As Name
    
    For Each nm In ThisWorkbook.Names
        Debug.Print nm.Name
    Next nm
    
End Sub

Sub DeleteAllNames()

    Dim nm As Name
    
    For Each nm In ThisWorkbook.Names
        Debug.Print nm.Name & " is deleted!"
        nm.Delete
    Next nm
    
End Sub

Public Sub DeleteCommentInSelection()
    
    If SET_IN_PRODUCTION Then On Error GoTo DeleteCommentInSelection_Error

    Dim myCell As Range
    
    For Each myCell In Selection
        myCell.ClearComments
    Next myCell
    
    On Error GoTo 0
    Exit Sub

DeleteCommentInSelection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DeleteCommentInSelection of Sub mod_StandardSubs"

End Sub

Public Sub SelectMeA1RangeEverywhere()

    If SET_IN_PRODUCTION Then On Error GoTo SelectMeA1RangeEverywhere_Error

    Dim wks As Worksheet

    For Each wks In ThisWorkbook.Worksheets
        If wks.Visible = xlSheetVisible Then
            wks.Activate
            wks.Cells(1, 1).Select
        End If
    Next
    
    Worksheets(1).Select

    On Error GoTo 0
    Exit Sub

SelectMeA1RangeEverywhere_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SelectMeA1RangeEverywhere of Sub mod_StandardSubs"

End Sub

Sub HideShowComments(Optional showComments As Boolean = False, _
                            Optional myRange As Range = Nothing)
    
    Dim myCell    As Range
    
    If SET_IN_PRODUCTION Then On Error GoTo HideShowComments_Error
    If myRange Is Nothing Then Set myRange = Range("A1:AO1000")
        
    For Each myCell In myRange
        If Not myCell.Comment Is Nothing Then
            myCell.Comment.Visible = showComments
        End If
    Next myCell

    On Error GoTo 0
    Exit Sub

HideShowComments_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HideShowComments of Sub mod_StandardSubs"

End Sub

Public Sub ResetAndUnlock()
    
    If Not IsValueInArray(Environ("Username"), ADMINS, True) Then
        Debug.Print "no"
        Exit Sub
    End If

    UnhideAll 'UnprotectAll is included
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"", true)"
    ActiveWindow.DisplayHeadings = True
    Application.DisplayFormulaBar = True
    Debug.Print "Done."

    EnableMySaves

End Sub

Public Sub EnableMySaves()

    Application.OnKey "%{F11}"
    Application.OnKey "^c"
    Application.OnKey "^C"
    Application.OnKey "^v"
    Application.OnKey "^V"
    Application.OnKey "^x"
    Application.OnKey "^X"
    Application.OnKey "^w"
    Application.OnKey "^W"
    Application.OnKey "^e"
    Application.OnKey "^E"

End Sub

Public Sub DisabledCombination()
    'This is the disabled combination for Application.OnKey
End Sub

Public Sub DisableShortcutsAndSaves()

    Application.OnKey "^c", "DisabledCombination"
    Application.OnKey "^C", "DisabledCombination"
    Application.OnKey "^v", "DisabledCombination"
    Application.OnKey "^V", "DisabledCombination"
    Application.OnKey "^x", "DisabledCombination"
    Application.OnKey "^X", "DisabledCombination"
    Application.OnKey "^w", "DisabledCombination"
    Application.OnKey "^W", "DisabledCombination"
    
    Application.OnKey "^e", "ShowMainForm"
    Application.OnKey "^E", "ShowMainForm"
    
End Sub
