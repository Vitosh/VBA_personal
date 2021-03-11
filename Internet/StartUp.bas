Attribute VB_Name = "StartUp"
Option Explicit

Public Sub Main()

    If IN_PRODUCTION Then On Error GoTo Main_Error
        
    CleanWorksheets
    Dim keyword As String: keyword = GetNextKeyWord
    
    While keyword <> ""
        
        Dim appIE As Object
        Set appIE = CreateObject("InternetExplorer.Application")
        LogMe keyword
        Dim nextPageExists As Boolean: nextPageExists = True
        Dim i As Long: i = 1
        Dim firstRow As Long: firstRow = lastRow(tblRawData.Name) + 1
        
        While nextPageExists
        
            WaitSomeMilliseconds
            Navigate i, appIE, keyword
            nextPageExists = PageWithResultsExists(appIE, keyword)
            If nextPageExists Then WriteToExcel appIE, keyword
            i = i + 1
            
        Wend
        
        LogMe Time, keyword, "RawDataToStructured"
        RawDataToStructured keyword, firstRow
        keyword = GetNextKeyWord
        WaitSomeMilliseconds 4000
        appIE.Quit
        
    Wend
    
    FixWorksheets
    WriteFormulas
    
    LogMe "Program has ended!"
    
    On Error GoTo 0
    Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main, line " & Erl & "."

End Sub
