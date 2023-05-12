VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ShowHideLL 
   Caption         =   "ShowNameApps"
   ClientHeight    =   4812
   ClientLeft      =   -20
   ClientTop       =   -80
   ClientWidth     =   9920.001
   OleObjectBlob   =   "F_ShowHideLL.frx":0000
   StartUpPosition =   1  'Propri�taireCentre
End
Attribute VB_Name = "F_ShowHideLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






























































































































































































































































Option Explicit
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"

Private showHideObject As ILLShowHide

Private Sub LST_LLVarNames_Click()
    showHideObject.UpdateVisibilityStatus LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Show_Click()
    showHideObject.ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    showHideObject.ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub CMD_Back_Click()
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim trads As ITranslation
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim dictsh As Worksheet
    
    'Initialize the showHideObject
    Set sh = ActiveSheet
    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set dictsh = ThisWorkbook.Worksheets(DICTSHEET)
    Set dict = LLDictionary.Create(dictsh, 1, 1)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set trads = lltrads.TransObject()
    Set showHideObject = LLShowHide.Create(trads, dict, sh)

    'Manage language
    Me.Caption = TranslateLLMsg(Me.Name)

    TranslateForm Me

    Me.width = 500
    Me.height = 350
End Sub









































































































































































































































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

