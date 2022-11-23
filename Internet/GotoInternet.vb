Public Sub Clicked(Optional b_logo As Boolean = False)

    Dim ie                  As Object
    Dim s_WebSites()        As Variant
    
   On Error GoTo Clicked_Error

    If b_logo Then
    
        s_WebSites = Array("https://www.facebook.com", _
                               "https://plus.google.com", _
                               "http://www.youtube.com")
    Else
        s_WebSites = Array("http://www.hoai.de/online/hoai_rechner")
    End If
     
'    s_WebSites = Array("https://goo.gl/c3Gzqi", _
'                        "https://goo.gl/JKvYR6", _
'                        "https://goo.gl/eLuMFN", _
'                        "https://goo.gl/r2OMeQ")
                
    Set ie = CreateObject("Internetexplorer.Application")
    ie.Visible = True
    ie.Navigate s_WebSites(make_random(0, UBound(s_WebSites)))

   Exit Sub

   On Error GoTo 0
   Exit Sub

Clicked_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Clicked of Module mod_main"
    
End Sub
            

Public Function CheckUrlExists(url) As Boolean
        
    On Error GoTo CheckUrlExists_Error
    
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
 
    xmlhttp.Open "HEAD", url, False
    xmlhttp.send
    
    If xmlhttp.Status = 200 Then
        CheckUrlExists = True
    Else
        CheckUrlExists = False
    End If
    
    Exit Function
    
CheckUrlExists_Error:
    CheckUrlExists = False
    
End Function
