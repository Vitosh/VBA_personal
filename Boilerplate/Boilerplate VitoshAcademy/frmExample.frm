VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmExample 
   Caption         =   "UserForm1"
   ClientHeight    =   2016
   ClientLeft      =   0
   ClientTop       =   204
   ClientWidth     =   3000
   OleObjectBlob   =   "frmExample.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event OnRunReport()
Public Event OnExit()

Public Property Get InformationText() As String

    InformationText = lblInfo.Caption

End Property

Public Property Let InformationText(ByVal value As String)

    lblInfo.Caption = value

End Property

Public Property Get InformationCaption() As String

    InformationCaption = Caption

End Property

Public Property Let InformationCaption(ByVal value As String)

    Caption = value

End Property

Private Sub btnRun_Click()

    RaiseEvent OnRunReport

End Sub

Private Sub btnExit_Click()

    RaiseEvent OnExit

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Hide
    End If

End Sub
