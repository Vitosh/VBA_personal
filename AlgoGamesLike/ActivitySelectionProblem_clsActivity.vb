Private pName       As String
Private pStartTime  As Long
Private pEndTime    As Long

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get StartTime() As Long
    StartTime = pStartTime
End Property

Public Property Let StartTime(value As Long)
    pStartTime = value
End Property

Public Property Get Endtime() As Long
    Endtime = pEndTime
End Property

Public Property Let Endtime(value As Long)
    pEndTime = value
End Property
