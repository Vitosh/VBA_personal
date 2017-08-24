Option Explicit

Public Sub Main()

    Dim objFso              As Object
    Dim objFol              As Object
    Dim objFil              As Object
    
    Dim objWb               As Workbook
    Dim objWs               As Worksheet
    
    Dim lngCounter          As Long
    Dim strNameToDelete     As String: strNameToDelete = UCase(tblMAin.Cells(1, 1))
    Dim strNameDeleted      As String
    
    Call OnStart
    
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFol = objFso.getfolder(ThisWorkbook.Path)
    strTextSummary = Now & vbCrLf
    
    Application.StatusBar = "Running ..."
    
    For Each objFil In objFol.Files
        If ((Not InStr(1, objFil.Name, "$") > 1) And _
            (Not InStr(1, objFil.Name, "~") > 1) And _
            (objFil.Name <> ThisWorkbook.Name) And _
            InStr(1, objFil.Name, "xls") > 1) Then
            
            Set objWb = Workbooks.Open(objFil.Path)
            Application.StatusBar = objFil.Name
            
            For lngCounter = objWb.Worksheets.Count To 1 Step -1
                If UCase(Left(objWb.Worksheets(lngCounter).Name, Len(strNameToDelete))) = strNameToDelete Then
                    strNameDeleted = objWb.Worksheets(lngCounter).Name
                    objWb.Worksheets(lngCounter).Delete
                    strTextSummary = strTextSummary & objWb.Name & vbCrLf & vbTab & strNameDeleted & vbCrLf
                End If
            Next lngCounter
            
            objWb.Close True
        
        End If
    Next objFil
    
    CreateLogFile
    Call OnEnd
        
End Sub
