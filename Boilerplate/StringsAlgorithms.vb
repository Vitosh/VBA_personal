Public Function StringBetween2Strings(ByVal myText As String, _
                        ByVal lookBefore As String, _
                        ByVal repetition As Long, _
                        Optional ByVal lookAfter As String = "</") _
                        As String
    
    On Error GoTo StringBetween2Strings_Error
    
    Dim i As Long: i = 1
    Dim startPosition As Long
    Dim endPosition As Long
    
    While repetition > 1
        i = InStr(i, myText, lookBefore, vbTextCompare)
        myText = Right(myText, Len(myText) - i)
        repetition = repetition - 1
    Wend
    
    startPosition = InStr(1, myText, lookBefore) + Len(lookBefore)
    endPosition = InStr(startPosition, myText, lookAfter, vbTextCompare)
    StringBetween2Strings = Mid(myText, startPosition, endPosition - startPosition)
    
    Exit Function
    
StringBetween2Strings_Error:
    StringBetween2Strings = -1

End Function

Sub TestingLocateXmlData()
    
    Dim xmlA As String
    xmlA = "<FootballInfo><row><ID>1</ID><FirstName>Peter</FirstName><LastName>The Keeper</LastName><Club name =NorthClub><ClubCoach>Pesho</ClubCoach><ClubManager>Partan</ClubManager><ClubEstablishedOn>1994</ClubEstablishedOn></Club><CityID>1</CityID></row><row name=Row2><ID>2</ID><FirstName>Ivan</FirstName><LastName>Mitov</LastName><Club name = EastClub><ClubCoach>Gosho</ClubCoach><ClubManager>Goshan</ClubManager><ClubEstablishedOn>1889</ClubEstablishedOn></Club><CityID>2</CityID></row>/FootballInfo>"
     
    Debug.Print StringBetween2Strings(xmlA, "<FirstName>", 1)   'Peter
    Debug.Print StringBetween2Strings(xmlA, "<LastName>", 1)    'The Keeper

    Debug.Print StringBetween2Strings(xmlA, "<ClubEstablishedOn>", 1)   '1994
    Debug.Print StringBetween2Strings(xmlA, "<ClubEstablishedOn>", 2)   '1889

End Sub