VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ShowHideLL 
   Caption         =   "ShowNameApps"
   ClientHeight    =   7410
   ClientLeft      =   -30
   ClientTop       =   -150
   ClientWidth     =   13755
   OleObjectBlob   =   "F_ShowHideLL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F_ShowHideLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Events show/hide in the linelist")

Option Explicit

Private Sub LST_LLVarNames_Click()
    ClickListShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Show_Click()
   ClickOptionsShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHide Me.LST_LLVarNames.ListIndex
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub
