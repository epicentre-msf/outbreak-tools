VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ShowHideLL 
   Caption         =   "ShowNameApps"
   ClientHeight    =   4812
   ClientLeft      =   -12
   ClientTop       =   -84
   ClientWidth     =   9912.001
   OleObjectBlob   =   "F_ShowHideLL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ShowHideLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





























































































































































































































































Option Explicit

Private Sub CMD_Back_Click()
    Me.Hide
End Sub

Private Sub LST_LLVarNames_Click()
    UpdateVisibilityStatus LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Show_Click()
    ShowHideLogic LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    ShowHideLogic LST_LLVarNames.ListIndex
End Sub

Private Sub UserForm_Initialize()
    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    Call TranslateForm(Me)

    Me.width = 450
    Me.height = 400

End Sub

