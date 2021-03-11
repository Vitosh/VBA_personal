VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoValueConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TypeMappings
  OptionAllStrings As Boolean
  OptionMapGuidString As Boolean
  StringDateFormat As String
  BooleanMap As ADODB.DataTypeEnum
  StringMap As ADODB.DataTypeEnum
  GuidMap As ADODB.DataTypeEnum
  DateMap As ADODB.DataTypeEnum
  ByteMap As ADODB.DataTypeEnum
  IntegerMap As ADODB.DataTypeEnum
  LongMap As ADODB.DataTypeEnum
  DoubleMap As ADODB.DataTypeEnum
  SingleMap As ADODB.DataTypeEnum
  CurrencyMap As ADODB.DataTypeEnum
End Type

Private mappings As TypeMappings
Option Explicit

Private Sub Class_Initialize()

  mappings.OptionAllStrings = False
  mappings.OptionMapGuidString = True
  mappings.StringDateFormat = "yyyy-MM-dd"

  mappings.BooleanMap = adBoolean
  mappings.ByteMap = adInteger
  mappings.CurrencyMap = adCurrency
  mappings.DateMap = adDate
  mappings.DoubleMap = adDouble
  mappings.GuidMap = adGUID
  mappings.IntegerMap = adInteger
  mappings.LongMap = adInteger
  mappings.SingleMap = adSingle
  mappings.StringMap = adVarChar

End Sub

Public Property Get OptionAllStrings() As Boolean
  OptionAllStrings = mappings.OptionAllStrings
End Property

Public Property Let OptionAllStrings(ByVal value As Boolean)
  mappings.OptionAllStrings = value
End Property

Public Property Get OptionMapGuidStrings() As Boolean
  OptionMapGuidStrings = mappings.OptionMapGuidString
End Property

Public Property Let OptionMapGuidStrings(ByVal value As Boolean)
  mappings.OptionMapGuidString = value
End Property

Public Property Get StringDateFormat() As String
  StringDateFormat = mappings.StringDateFormat
End Property

Public Property Let StringDateFormat(ByVal value As String)
  mappings.StringDateFormat = value
End Property

Public Property Get BooleanMapping() As ADODB.DataTypeEnum
  BooleanMapping = mappings.BooleanMap
End Property

Public Property Let BooleanMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.BooleanMap = value
End Property

Public Property Get ByteMapping() As ADODB.DataTypeEnum
  ByteMapping = mappings.ByteMap
End Property

Public Property Let ByteMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.ByteMap = value
End Property

Public Property Get CurrencyMapping() As ADODB.DataTypeEnum
  CurrencyMapping = mappings.CurrencyMap
End Property

Public Property Let CurrencyMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.CurrencyMap = value
End Property

Public Property Get DateMapping() As ADODB.DataTypeEnum
  DateMapping = mappings.DateMap
End Property

Public Property Let DateMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.DateMap = value
End Property

Public Property Get DoubleMapping() As ADODB.DataTypeEnum
  DoubleMapping = mappings.DoubleMap
End Property

Public Property Let DoubleMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.DoubleMap = value
End Property

Public Property Get GuidMapping() As ADODB.DataTypeEnum
  GuidMapping = mappings.GuidMap
End Property

Public Property Let GuidMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.GuidMap = value
End Property

Public Property Get IntegerMapping() As ADODB.DataTypeEnum
  IntegerMapping = mappings.IntegerMap
End Property

Public Property Let IntegerMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.IntegerMap = value
End Property

Public Property Get LongMapping() As ADODB.DataTypeEnum
  LongMapping = mappings.LongMap
End Property

Public Property Let LongMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.LongMap = value
End Property

Public Property Get SingleMapping() As ADODB.DataTypeEnum
  SingleMapping = mappings.SingleMap
End Property

Public Property Let SingleMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.SingleMap = value
End Property

Public Property Get StringMapping() As ADODB.DataTypeEnum
  StringMapping = mappings.StringMap
End Property

Public Property Let StringMapping(ByVal value As ADODB.DataTypeEnum)
  mappings.StringMap = value
End Property

Public Function ToNamedParameter(ByVal name As String, ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  Dim result As ADODB.Parameter
  Set result = CallByName(Me, "To" & TypeName(value) & "Parameter", VbMethod, value, direction)

  result.name = name
  Set ToNamedParameter = result

End Function

Public Function ToStringParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  Dim stringValue As String
  stringValue = CStr(value)

  If Not mappings.OptionAllStrings Then
    If IsGuidString(stringValue) Then ' split on 2 conditions for performance: evaluating IsGuidString uses regular expressions
      Set ToStringParameter = ToGuidParameter(value, direction)
      Exit Function
    End If
  End If

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.StringMap
    .direction = direction
    .Size = Len(stringValue)
    .value = stringValue
  End With

  Set ToStringParameter = result

End Function

Public Function ToGuidParameter(ByVal value As String, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToGuidParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.GuidMap
    .direction = direction
    .value = value
  End With

  Set ToGuidParameter = result

End Function

Private Function IsGuidString(ByVal value As String) As Boolean

  Dim regex As New RegExp
  regex.Pattern = "\b[A-F0-9]{8}(?:-[A-F0-9]{4}){3}-[A-F0-9]{12}\b"

  Dim matches As MatchCollection
  Set matches = regex.Execute(UCase$(value))

  IsGuidString = matches.Count <> 0

  Set regex = Nothing
  Set matches = Nothing

End Function

Public Function ToIntegerParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToIntegerParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim integerValue As Long
  integerValue = CLng(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.IntegerMap
    .direction = direction
    .value = integerValue
  End With

  Set ToIntegerParameter = result

End Function

Public Function ToByteParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToByteParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim byteValue As Byte
  byteValue = CByte(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.ByteMap
    .direction = direction
    .value = byteValue
  End With

  Set ToByteParameter = result

End Function

Public Function ToLongParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToLongParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim longValue As Long
  longValue = CLng(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.LongMap
    .direction = direction
    .value = longValue
  End With

  Set ToLongParameter = result

End Function

Public Function ToDoubleParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToDoubleParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim doubleValue As Double
  doubleValue = CDbl(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.DoubleMap
    .direction = direction
    .value = doubleValue
  End With

  Set ToDoubleParameter = result

End Function

Public Function ToSingleParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToSingleParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim singleValue As Single
  singleValue = CSng(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.SingleMap
    .direction = direction
    .value = singleValue
  End With

  Set ToSingleParameter = result

End Function

Public Function ToCurrencyParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToCurrencyParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim currencyValue As Currency
  currencyValue = CCur(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.CurrencyMap
    .direction = direction
    .value = currencyValue
  End With

  Set ToCurrencyParameter = result

End Function

Public Function ToBooleanParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToBooleanParameter = ToStringParameter(value, direction)
    Exit Function
  End If

  Dim boolValue As Boolean
  boolValue = CBool(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.BooleanMap
    .direction = direction
    .value = boolValue
  End With

  Set ToBooleanParameter = result

End Function

Public Function ToDateParameter(ByVal value As Variant, ByVal direction As ADODB.ParameterDirectionEnum) As ADODB.Parameter

  If mappings.OptionAllStrings Then
    Set ToDateParameter = ToStringParameter(Format$(value, mappings.StringDateFormat), direction)
    Exit Function
  End If

  Dim dateValue As Date
  dateValue = CDate(value)

  Dim result As ADODB.Parameter
  Set result = New ADODB.Parameter
  With result
    .Type = mappings.DateMap
    .direction = direction
    .value = dateValue
  End With

  Set ToDateParameter = result

End Function
