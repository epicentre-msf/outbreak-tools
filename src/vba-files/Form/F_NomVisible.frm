VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_NomVisible 
   Caption         =   "ShowNameApps"
   ClientHeight    =   5880
   ClientLeft      =   -36
   ClientTop       =   -168
   ClientWidth     =   8304.001
   OleObjectBlob   =   "F_NomVisible.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_NomVisible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub CMD_Fermer_Click()

    F_NomVisible.Hide

End Sub

Private Sub LST_NomChamp_Click()

    If Not bLockActu Then
        Call IsVisibleDataName(LST_NomChamp.value)
    End If

End Sub

Private Sub OPT_Affiche_Click()

    If Not bLockActu Then
        Call ShowDataCol(LST_NomChamp.ListIndex)
    End If

End Sub

Private Sub OPT_Masque_Click()

    If Not bLockActu Then
        Call HideDataCol(LST_NomChamp.ListIndex)
    End If

End Sub

