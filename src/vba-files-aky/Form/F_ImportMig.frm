VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ImportMig 
   Caption         =   "Import For Migration"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2955
   OleObjectBlob   =   "F_ImportMig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ImportMig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub CMD_ClearGeo_Click()
    Call ClearHistoricGeobase
End Sub

Private Sub CMD_ImportData_Click()
    Call ImportMigrationData
End Sub

Private Sub CMD_ImportGeo_Click()
    Call ImportGeobase
End Sub

Private Sub CMD_ImportGeoHistoric_Click()
    Call ImportHistoricGeobase
End Sub

Private Sub CMD_ImportMigQuit_Click()
    Me.Hide
End Sub

