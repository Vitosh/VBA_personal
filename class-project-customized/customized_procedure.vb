Public Sub PrintProperties(my_object As Object)
    'Tools - References - TypeLib Information
    
    Dim mi                      As TLI.MemberInfo
    Dim ti                      As TLI.TypeInfo
    Dim t                       As TLI.TLIApplication
    
    Set t = New TLI.TLIApplication
    
    Set ti = t.InterfaceInfoFromObject(my_object)
    
    Debug.Print "***********************"
    
    For Each mi In ti.Members
            '0 is for GET Properties,
            '1 is for LET Properties
            'Change accordingly
            If mi.ReturnType.PointerLevel = 0 Then
                Debug.Print mi.name & vbCrLf; CallByName(my_object, mi.name, VbGet) & vbCrLf
            End If
    Next mi
    
    Debug.Print "***********************"
    
    Set my_object = Nothing

    
End Sub
