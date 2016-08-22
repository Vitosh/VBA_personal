Option Explicit

Sub SQL()

Dim cn      As Object
Dim rs      As Object
Dim strfile As String
Dim strCon  As String
Dim strSQL  As String

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

strfile = ThisWorkbook.FullName
strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strfile & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Open strCon

strSQL = "SELECT * FROM [Tabelle1$A1:C5]"

rs.Open strSQL, cn

Debug.Print rs.GetString

Set cn = Nothing
Set rs = Nothing

End Sub
