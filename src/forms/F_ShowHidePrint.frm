VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ShowHidePrint 
   Caption         =   "ShowNameApps"
   ClientHeight    =   6372
   ClientLeft      =   -12
   ClientTop       =   -84
   ClientWidth     =   9780.001
   OleObjectBlob   =   "F_ShowHidePrint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ShowHidePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





























































































































































































































































Option Explicit


Private Sub LST_NomChamp_Click()
    'UpdateVisibilityStatus LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Affiche_Click()
    'ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Masque_Click()
    'ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub CMD_PrintBack_Click()
    Me.Hide
End Sub



Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    TranslateForm Me

    Me.width = 500
    Me.height = 350

End Sub

