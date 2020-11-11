Option Explicit

Sub myPathForFolder()
    Debug.Print GetFolder(Environ("USERPROFILE"))
End Sub

Function GetFolder(Optional InitialLocation As String) As String

    On Error GoTo GetFolder_Error

    Dim FolderDialog        As FileDialog
    Dim SelectedFolder      As String

    If Len(InitialLocation) = 0 Then InitialLocation = ThisWorkbook.Path

    Set FolderDialog = Excel.Application.FileDialog(msoFileDialogFolderPicker)

    With FolderDialog
        .Title = "My Title For Dialog"
        .AllowMultiSelect = False
        .InitialFileName = InitialLocation
        If .Show <> -1 Then GoTo GetFolder_Error
        SelectedFolder = .SelectedItems(1)
    End With

    GetFolder = SelectedFolder

    On Error GoTo 0
    Exit Function

GetFolder_Error:

    Debug.Print "Error " & Err.Number & " (" & Err.Description & ")

End Function

'---------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------
'Taken from http://www.cpearson.com/excel/browsefolder.aspx

Private Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As Long, ByVal pszBuffer As String) As Long
Private Declare Function SHBrowseForFolderA Lib "shell32.dll" (lpBrowseInfo As BROWSEINFO) As Long
Private Const MAX_PATH = 260

Function str_BrowseFolder(Optional ByVal DialogTitle As String) As String

    On Error GoTo str_BrowseFolder_Error

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' BrowseFolder
    ' This displays the standard Windows Browse Folder dialog. It returns
    ' the complete path name of the selected folder or vbNullString if the
    ' user cancelled.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.EnableCancelKey = xlDisabled

    If DialogTitle = vbNullString Then
        DialogTitle = "Select A Folder"
    End If

    Dim uBrowseInfo     As BROWSEINFO
    Dim szBuffer        As String
    Dim lID             As Long
    Dim lRet            As Long

    With uBrowseInfo
        .hOwner = 0
        .pidlRoot = 0
        .pszDisplayName = String$(MAX_PATH, vbNullChar)
        .lpszINSTRUCTIONS = DialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS    ' + BIF_USENEWUI
        .lpfn = 0
    End With
    
    szBuffer = String$(MAX_PATH, vbNullChar)
    lID = SHBrowseForFolderA(uBrowseInfo)

    If lID Then
        ''' Retrieve the path string.
        lRet = SHGetPathFromIDListA(lID, szBuffer)
        If lRet Then
            str_BrowseFolder = Left$(szBuffer, InStr(szBuffer, vbNullChar) - 1)
        End If
    End If
    
    Application.EnableCancelKey = xlInterrupt

    On Error GoTo 0
    Exit Function

str_BrowseFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure str_BrowseFolder of Function mod_Browse"

End Function

            
Public Function FolderIsEmpty(myPath As String) As Boolean
    'Checks whether folder is empty    
    FolderIsEmpty = CBool(Dir(myPath & "*.*") = "")
    
End Function
