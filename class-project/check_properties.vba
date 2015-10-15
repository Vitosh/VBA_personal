Public Sub GetInformationPrinted()
'Tools - References - TypeLib Information

    Dim k                       As cls_arrCalendarSettings
    Dim mi                      As TLI.MemberInfo
    Dim ti                      As TLI.TypeInfo
    Dim t                       As TLI.TLIApplication
    Dim b_show                  As Boolean
    
    Set k = New cls_arrCalendarSettings
    
    k.TopRow = 10
    k.BottomRow = 15
    k.LeftCol = 3
    k.RightCol = 10
    k.SonstigesProBA = 1000.12
    k.VerhaltnisBaukostenToPlanerkosten = 0.35
    k.Vertriebsstart = Now()
    k.Vertriebsstart_Col = 50

    'Now printing all
    Set t = New TLI.TLIApplication
    
    Set ti = t.InterfaceInfoFromObject(k)

    For Each mi In ti.Members
            '0 is for LET Properties, 1 is for LET Properties
            'Change accordingly
            If mi.ReturnType.PointerLevel = 0 Then
                Debug.Print mi.name & vbCrLf; CallByName(k, mi.name, VbGet) & vbCrLf
            End If
    Next mi
    
    Set k = Nothing

End Sub
