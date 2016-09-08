Option Compare Database
Option Explicit

Public Sub TestTheseQueries()

    Dim rst                 As Recordset
    Dim dbeError            As Error
    
    On Error GoTo TestTheseQueries_Error
      
    Set rst = CurrentDb.OpenRecordset("SELECT TOP 1 frs_invoice.paid_amount_net FROM frs_invoice;")
    Debug.Print [rst]![paid_amount_net]
    Set rst = Nothing
    
    Exit Sub

TestTheseQueries_Error:
    
    For Each dbeError In DBEngine.Errors
        Debug.Print dbeError.Number & "->"; dbeError.Description
    Next dbeError
    
    Set rst = Nothing
    
End Sub
