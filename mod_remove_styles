Option Explicit

Sub RemoveTheStyles()

    Dim s       As Style
    Dim i       As Long
    Dim c       As Long
    
    If ActiveWorkbook.MultiUserEditing Then
        If MsgBox("You cannot remove Styles in a Shared workbook." & vbCr & vbCr & _
                  "Do you want to unshare the workbook?", vbYesNo + vbInformation) = vbYes Then
            ActiveWorkbook.ExclusiveAccess
            If Err.Description = "Application-defined or object-defined error" Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    c = ActiveWorkbook.Styles.Count
    Application.ScreenUpdating = False

    For i = c To 1 Step -1
    
        If i Mod 600 = 0 Then DoEvents
        Set s = ActiveWorkbook.Styles(i)
        Application.StatusBar = "Deleting " & c - i + 1 & " of " & c & " " & s.Name
        
        If Not s.BuiltIn Then
            s.Delete
        End If
    Next
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub
