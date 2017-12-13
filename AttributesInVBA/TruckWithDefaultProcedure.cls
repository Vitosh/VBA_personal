VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TruckWithDefaultProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sModel As String
Private m_Price As Currency

Private Sub Class_Initialize()
    
    Model = "Some truck model"
    Price = 250
    
End Sub

Public Property Get Model() As String

    Model = m_sModel

End Property

Public Property Let Model(ByVal sNewValue As String)

    m_sModel = sNewValue

End Property


Public Property Get Price() As Currency

    Price = m_Price

End Property

Public Property Let Price(ByVal NewValue As Currency)

    m_Price = NewValue

End Property

Public Function IncreasePriceWith10()
Attribute IncreasePriceWith10.VB_Description = "Increases the price with 10. It is the default."
Attribute IncreasePriceWith10.VB_UserMemId = 0
        Price = Price + 10
        Debug.Print "The price is now " & Price
End Function
