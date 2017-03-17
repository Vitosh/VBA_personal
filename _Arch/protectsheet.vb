Option Explicit
Sub main()
    
    'This protects the code only
    tbl_main.Protect UserInterfaceOnly:=True

End Sub


Public Sub UnprotectAll()

    Dim i As Long
    
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
        ActiveWorkbook.Worksheets(i).Unprotect Password:=s_CONST
    Next

End Sub

Public Sub UnhideAll()
        
    Dim Sheet As Worksheet
    
    For Each Sheet In ThisWorkbook.Sheets
        Sheet.Visible = xlSheetVisible
    Next Sheet
        
End Sub
