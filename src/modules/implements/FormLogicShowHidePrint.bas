Attribute VB_Name = "FormLogicShowHidePrint"
Attribute VB_Description = "Events show/hide in the printed linelist"
Option Explicit


Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"

Private showHideObject As ILLShowHide

Private Sub LST_PrintNames_Click()
    showHideObject.UpdateVisibilityStatus LST_PrintNames.ListIndex
End Sub

Private Sub OPT_PrintShowHoriz_Click()
    showHideObject.ShowHideLogic LST_PrintNames.ListIndex
End Sub

Private Sub OPT_PrintShowVerti_Click()
    showHideObject.ShowHideLogic LST_PrintNames.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    showHideObject.ShowHideLogic LST_PrintNames.ListIndex
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
    Set dictsh = ThisWorkbook.Worksheets(DictSheet)
    Set dict = LLdictionary.Create(dictsh, 1, 1)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set trads = lltrads.TransObject()

    Set showHideObject = LLShowHide.Create(trads, dict, sh)
    Me.Caption = trads.TranslatedValue(Me.Name)

    Set trads = lltrads.TransObject(TranslationOfForms)
    showHideObject.Load


    'Manage language
    trads.TranslateForm Me
End Sub
