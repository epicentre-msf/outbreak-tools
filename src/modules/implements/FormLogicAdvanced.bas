
Attribute VB_Name = "FormLogicAdvanced"
Attribute VB_Description = "Form implementation of advanced"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
Option Explicit

Private tradform As ITranslation   'Translation of forms
Private geoObj As ILLGeo 'Geo object
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const GEOSHEET As String = "Geo"

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim geosh As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set lltranssh = wb.Worksheets(LLSHEET)
    Set dicttranssh = wb.Worksheets(TRADSHEET)
    SEt geosh = wb.Worksheets(GEOSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set geoObj = LLGeo.Create(geosh)
End Sub


Private Sub CMD_ClearData_Click()
    Call ControlClearData
End Sub

Private Sub CMD_ClearGeo_Click()
    Call ClearHistoricGeobase
End Sub

Private Sub CMD_ExportData_Click()
    F_Advanced.Hide
    ClickExportMigration
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
