Attribute VB_Name = "FormLogicExportMigration"
Attribute VB_Description = "Form implementation of Exports for Migration"

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


Private Sub CMD_ExportMig_Click()
    Call ExportForMigration
End Sub

Private Sub CMD_ExportMigQuit_Click()
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    'Manage language
    Me.Caption = tradform.TranslatedValue(Me.Name)

    tradform.TranslateForm  Me

    Me.width = 200
    Me.height = 300
End Sub