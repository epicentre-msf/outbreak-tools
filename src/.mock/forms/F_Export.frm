VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Export 
   Caption         =   "Export"
   ClientHeight    =   5928
   ClientLeft      =   12
   ClientTop       =   -96
   ClientWidth     =   4512
   OleObjectBlob   =   "F_Export.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




































































































































































































































































Option Explicit

Private Sub CMD_Export2_Click()

    Call export(2)

End Sub

Private Sub CMD_Export1_Click()

    Call export(1)

End Sub

Private Sub CMD_Export3_Click()

    Call export(3)

End Sub

Private Sub CMD_Export4_Click()

    Call export(4)

End Sub

Private Sub CMD_Export5_Click()

    Call export(5)

End Sub

Private Sub CMD_NouvCle_Click()

    Call NewKey

End Sub

Private Sub CMD_Retour_Click()

    F_Export.Hide

End Sub

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 200
    Me.height = 400

End Sub
