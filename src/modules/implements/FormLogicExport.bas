Attribute VB_Name = "FormLogicExport"
Attribute VB_Description = "Form implementation of Exports"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable

Option Explicit

Private tradform As ITranslation   'Translation of forms
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet


    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub CMD_Export2_Click()

    Call export(2)

End Sub

Private Sub CMD_Export1_Click()

    Call export(1)

End Sub

Private Sub CMD_Export3_Click()

    Call export(3)

End Sub

Private Sub CMD_Export4_Click()

    Call export(4)

End Sub

Private Sub CMD_Export5_Click()

    Call export(5)

End Sub

Private Sub CMD_NouvCle_Click()

    Call NewKey

End Sub

Private Sub CMD_Retour_Click()

    F_Export.Hide

End Sub

'Translate the form, add form sizes.
Private Sub UserForm_Initialize()
    'Manage language
    
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 200
    Me.height = 400
End Sub
