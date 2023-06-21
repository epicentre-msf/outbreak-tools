
Attribute VB_Name = "FormLogicAdvanced"
Attribute VB_Description = "Form implementation of advanced"

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


Private Sub CMD_ClearData_Click()
    Call ControlClearData
End Sub

Private Sub CMD_ClearGeo_Click()
    Call ClearHistoricGeobase
End Sub

Private Sub CMD_ExportData_Click()
    F_Advanced.Hide
    Call ClicExportMigration
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

Private Sub CMD_ImportMigRep_Click()
    Me.Hide
    ShowImportReport
End Sub

Private Sub UserForm_Initialize()

    'Manage language of the userform
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 250
    Me.height = 550
End Sub
