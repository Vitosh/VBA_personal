VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarWithDefaultProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sModel    As String
Private m_Price     As Currency

Private Sub Class_Initialize()
    
    Model = "Car with default property model"
    Price = 2000
    
End Sub

Public Property Get Model() As String

    Model = m_sModel

End Property

Public Property Let Model(ByVal sNewValue As String)

    m_sModel = sNewValue

End Property

Public Property Get Price() As Currency
Attribute Price.VB_Description = "Some nice description should be here."
Attribute Price.VB_UserMemId = 0

    Price = m_Price

End Property

Public Property Let Price(ByVal NewValue As Currency)

    m_Price = NewValue

End Property

Public Function ChangePrice(newPrice As Currency) As Currency
    
    Price = Price + newPrice
    ChangePrice = Price
    
End Function

