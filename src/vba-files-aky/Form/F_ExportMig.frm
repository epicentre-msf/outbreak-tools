VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ExportMig 
   Caption         =   "Export For Migration"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "F_ExportMig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ExportMig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





















































Option Explicit

Private Sub CMD_ExportMig_Click()
    Call ExportForMigration
End Sub


Private Sub CMD_ExportMigQuit_Click()
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.Width = 200
    Me.Height = 300

End Sub
