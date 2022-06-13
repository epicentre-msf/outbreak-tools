VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ImportRep 
   Caption         =   "Import Summary"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "F_ImportRep.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ImportRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

























Option Explicit

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.Width = 300
    Me.Height = 550

End Sub

Private Sub CMD_ImpRepQuit_Click()
    Me.Hide
End Sub
