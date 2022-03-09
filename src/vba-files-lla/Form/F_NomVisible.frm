VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_NomVisible 
   Caption         =   "ShowNameApps"
   ClientHeight    =   5880
   ClientLeft      =   -30
   ClientTop       =   -165
   ClientWidth     =   8295.001
   OleObjectBlob   =   "F_NomVisible.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_NomVisible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Option Explicit
Option Base 1

Private Sub CMD_Fermer_Click()
    F_NomVisible.Hide
    WriteVisibility
End Sub

Private Sub LST_NomChamp_Click()
    UpdateVisibilityStatus LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Affiche_Click()
    ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Masque_Click()
    ShowHideLogic LST_NomChamp.ListIndex
End Sub


