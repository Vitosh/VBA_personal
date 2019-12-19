VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmInfo 
   ClientHeight    =   1008
   ClientLeft      =   12
   ClientTop       =   84
   ClientWidth     =   2532
   OleObjectBlob   =   "frmInfo.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
        
    If PUB_STR_ERROR_REPORT Then
        Me.lb_information = CON_STR_INSTANCES_LOG
    End If
    
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Caption = CON_STR_APP_NAME
    End With

End Sub
