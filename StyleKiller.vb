Option Explicit

Sub StyleKiller()

    Dim myStyle                As Style
    Dim lngCounter              As Long
    
    For Each myStyle In ThisWorkbook.Styles        
        
        If Not myStyle.BuiltIn Then
            Debug.Print myStyle.name
            myStyle.Delete
            lngCounter = lngCounter + 1
        End If
    Next myStyle
    
    Debug.Print "Ende"
    Debug.Print "Deleted " & lngCounter
    
End Sub

'FANCY ONE:
'**************************************************************************************
Sub RemoveTheStyles()

    Dim style               As style
    Dim l_counter           As Long
    Dim l_total_number      As Long

    On Error Resume Next

    l_total_number = ActiveWorkbook.Styles.Count
    Application.ScreenUpdating = False

    For l_counter = l_total_number To 1 Step -1
    
        Set style = ActiveWorkbook.Styles(l_counter)
        
        If (l_counter Mod 500 = 0) Then
            DoEvents
            Application.StatusBar = "Deleting " & l_total_number - l_counter + 1 & " of " & l_total_number & " " & style.Name
        End If
        
        If Not style.BuiltIn Then style.Delete

    Next l_counter

    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print "READY!"
    
    On Error GoTo 0
End Sub


'https://support.microsoft.com/en-us/help/291321/how-to-programmatically-reset-a-workbook-to-default-styles
Sub RebuildDefaultStyles()

'The purpose of this macro is to remove all styles in the active
'workbook and rebuild the default styles.
'It rebuilds the default styles by merging them from a new workbook.

'Dimension variables.
   Dim MyBook As Workbook
   Dim tempBook As Workbook
   Dim CurStyle As Style

   'Set MyBook to the active workbook.
   Set MyBook = ActiveWorkbook
   On Error Resume Next
   'Delete all the styles in the workbook.
   For Each CurStyle In MyBook.Styles
      'If CurStyle.Name <> "Normal" Then CurStyle.Delete
      Select Case CurStyle.Name
         Case "20% - Accent1", "20% - Accent2", _
               "20% - Accent3", "20% - Accent4", "20% - Accent5", "20% - Accent6", _
               "40% - Accent1", "40% - Accent2", "40% - Accent3", "40% - Accent4", _
               "40% - Accent5", "40% - Accent6", "60% - Accent1", "60% - Accent2", _
               "60% - Accent3", "60% - Accent4", "60% - Accent5", "60% - Accent6", _
               "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", _
               "Bad", "Calculation", "Check Cell", "Comma", "Comma [0]", "Currency", _
               "Currency [0]", "Explanatory Text", "Good", "Heading 1", "Heading 2", _
               "Heading 3", "Heading 4", "Input", "Linked Cell", "Neutral", "Normal", _
               "Note", "Output", "Percent", "Title", "Total", "Warning Text"
            'Do nothing, these are the default styles
         Case Else
            CurStyle.Delete
      End Select

   Next CurStyle

   'Open a new workbook.
   Set tempBook = Workbooks.Add

   'Disable alerts so you may merge changes to the Normal style
   'from the new workbook.
   Application.DisplayAlerts = False

   'Merge styles from the new workbook into the existing workbook.
   MyBook.Styles.Merge Workbook:=tempBook

   'Enable alerts.
   Application.DisplayAlerts = True

   'Close the new workbook.
   tempBook.Close

End Sub
