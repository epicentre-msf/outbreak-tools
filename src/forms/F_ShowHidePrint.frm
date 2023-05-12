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

Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Name = "F_ShowHidePrint"


Option Explicit
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"

Private showHideObject As ILLShowHide

Private Sub LST_PrintNames_Click()
    showHideObject.UpdateVisibilityStatus LST_NomChamp.ListIndex
End Sub

Private Sub OPT_PrintShowHoriz_Click()
    showHideObject.ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub OPT_PrintShowVerti_Click()
    showHideObject.ShowHideLogic LST_NomChamp.ListIndex
End Sub

Private Sub OPT_Hide_Click()
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

    showHideObject.Load
    Me.Caption = trads.TranslatedValue(Me.Name)

    'Manage language
    trads.TranslateForm Me

    Me.width = 500
    Me.height = 350
End Sub