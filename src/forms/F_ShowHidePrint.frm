VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_ShowHidePrint 
   Caption         =   "ShowNameApps"
   ClientHeight    =   6372
   ClientLeft      =   -20
   ClientTop       =   -80
   ClientWidth     =   9780.001
   OleObjectBlob   =   "F_ShowHidePrint.frx":0000
   StartUpPosition =   1  'Propri�taireCentre
End
 = "F_ShowHidePrint"
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

Attribute VB_Name = "F_ShowHidePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

























































































































































































































































Option Explicit
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"

Private showHideObject As ILLShowHide

Private Sub LST_NomChamp_Click()
    showHideObject.UpdateVisibilityStatus LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Affiche_Click()
    showHideObject.ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Masque_Click()
    showHideObject.ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub CMD_PrintBack_Click()
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

