Option Explicit

Public Sub TestMe()

    Dim lngCounter          As Long
    Dim strURL              As String
    Dim IE                  As Object
    Dim colCurrent          As Object
    Dim link
    Dim colLinks            As Collection
    
    strURL = "vitoshacademy.com"
    Set IE = CreateObject("InternetExplorer.Application")
    Set colLinks = New Collection
        
    'IE.Visible = True
    IE.navigate strURL
    Application.Wait (Now() + TimeValue("0:00:2"))
    
    Set colCurrent = IE.Document.getElementsByTagName("a")
    For Each link In colCurrent
        'link.Click
        'Application.Wait (Now() + TimeValue("0:00:2"))
        If Not Contains(colLinks, link) Then colLinks.Add (link)
        Debug.Print link.href
        'Debug.Print link.textContent
        'Debug.Print link.OuterHTML
        'Debug.Print "-------------------"
    Next link
    
'    For Each link In colLinks
'        IE.navigate strURL
'        If Not Contains(colLinks, link) Then colLinks.Add (link)
'    Next link
    
    
    
    Stop
    IE.Quit
    
End Sub

Public Function Contains(col As Collection, key As Variant) As Boolean
    
    Dim var     As Variant
    
    For Each var In col
        If var = key Then
            Contains = True
            Exit Function
        End If
    Next var
    
    Contains = False
 
End Function
