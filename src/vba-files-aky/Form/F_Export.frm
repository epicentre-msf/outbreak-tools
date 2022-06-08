VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Export 
   Caption         =   "Export"
   ClientHeight    =   5415
   ClientLeft      =   -15
   ClientTop       =   -180
   ClientWidth     =   3765
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

    Call Export(2)

End Sub

Private Sub CMD_Export1_Click()

    Call Export(1)

End Sub

Private Sub CMD_Export3_Click()

    Call Export(3)

End Sub

Private Sub CMD_Export4_Click()

    Call Export(4)

End Sub

Private Sub CMD_Export5_Click()

    Call Export(5)

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

    Me.Width = 200
    Me.Height = 300

End Sub
