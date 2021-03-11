VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private converter As New AdoValueConverter

Public Property Get ParameterFactory() As AdoValueConverter
  Set ParameterFactory = converter
End Property

Public Function Execute(ByVal connection As ADODB.connection, ByVal sql As String, ParamArray parameterValues()) As ADODB.Recordset
'Returns a connected ADODB.Recordset that contains the results of the specified parameterized query.

  Dim parameters() As Variant
  parameters = parameterValues

  Set Execute = ExecuteInternal(connection, sql, parameters)

End Function

Public Function ExecuteNonQuery(ByVal connection As ADODB.connection, ByVal sql As String, ParamArray parameterValues()) As Boolean
'Returns a Boolean that indicates whether the specified parameterized SQL command (update, delete, etc.) executed without throwing an error.

  Dim parameters() As Variant
  parameters = parameterValues

  ExecuteNonQuery = ExecuteNonQueryInternal(connection, sql, parameters)

End Function

Public Function ExecuteStoredProc(ByVal connection As ADODB.connection, ByVal spName As String, ParamArray parameterValues()) As ADODB.Recordset
'Executes the specified parameterized stored procedure, passing specified parameter values.

  Dim parameters() As Variant
  parameters = parameterValues

  Set ExecuteStoredProc = ExecuteStoredProcInternal(connection, spName, parameters)

End Function

Public Function SelectSingleValue(ByVal connection As ADODB.connection, ByVal sql As String, ParamArray parameterValues()) As Variant
'Returns the value of the first field of the first record of the results of the specified parameterized SQL query.

  Dim parameters() As Variant
  parameters = parameterValues

  SelectSingleValue = SelectSingleValueInternal(connection, sql, parameters)

End Function

Private Function CreateCommand(ByVal connection As ADODB.connection, ByVal cmdType As ADODB.CommandTypeEnum, ByVal sql As String, parameterValues() As Variant) As ADODB.Command

  Dim cmd As ADODB.Command
  Set cmd = New ADODB.Command
  cmd.ActiveConnection = connection
  cmd.CommandType = cmdType
  cmd.CommandText = sql

  Dim i As Long
  Dim value As Variant

  For i = LBound(parameterValues) To UBound(parameterValues)
    value = parameterValues(i)
    If TypeName(value) <> "Variant()" Then cmd.parameters.Append ToSqlInputParameter(value)
  Next

  Set CreateCommand = cmd

End Function

Private Function ToSqlInputParameter(ByVal value As Variant) As ADODB.Parameter

  If IsObject(value) Then Err.Raise vbObjectError + 911, "SqlCommand.ToSqlInputParameter", "Invalid argument, parameter value cannot be an object."

  Dim result As ADODB.Parameter
  Set result = CallByName(converter, "To" & TypeName(value) & "Parameter", VbMethod, value, ADODB.ParameterDirectionEnum.adParamInput)

  Set ToSqlInputParameter = result

End Function

Private Function ExecuteInternal(ByVal connection As ADODB.connection, ByVal sql As String, parameterValues()) As ADODB.Recordset

  Dim cmd As ADODB.Command
  Set cmd = CreateCommand(connection, adCmdText, sql, parameterValues)

  Set ExecuteInternal = cmd.Execute

End Function

Private Function ExecuteNonQueryInternal(ByVal connection As ADODB.connection, ByVal sql As String, parameterValues()) As Boolean

  Dim cmd As ADODB.Command
  Set cmd = CreateCommand(connection, adCmdText, sql, parameterValues)

  Dim result As Boolean
  On Error Resume Next
  cmd.Execute
  result = (Err.Number = 0)
  On Error GoTo 0

  ExecuteNonQueryInternal = result

End Function

Private Function ExecuteStoredProcInternal(ByVal connection As ADODB.connection, ByVal spName As String, parameterValues()) As ADODB.Recordset

  Dim cmd As ADODB.Command
  Set cmd = CreateCommand(connection, adCmdStoredProc, spName, parameterValues)

  Set ExecuteStoredProcInternal = cmd.Execute

End Function

Private Function SelectSingleValueInternal(ByVal connection As ADODB.connection, ByVal sql As String, parameterValues()) As Variant

  Dim parameters() As Variant
  parameters = parameterValues

  Dim cmd As ADODB.Command
  Set cmd = CreateCommand(connection, adCmdText, sql, parameters)

  Dim rs As ADODB.Recordset
  Set rs = cmd.Execute

  Dim result As Variant
  If Not rs.BOF And Not rs.EOF Then result = rs.Fields(0).value

  rs.Close
  Set rs = Nothing

  SelectSingleValueInternal = result

End Function
