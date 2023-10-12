Attribute VB_Name = "FormLogicAdvanced"
Attribute VB_Description = "Form implementation of advanced form"


'@Folder("Form logics")
'@ModuleDescription("Form implementation of advanced form")
'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
Option Explicit

Private tradform As ITranslation   'Translation of forms
Private tradmess As ITranslation 'Translation of messages
Private currwb As Workbook 'Current workbook (used for creating import classes)
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const GEOSHEET As String = "Geo"


'Initialize translation of forms object
Private Sub Initialize()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet

    Set currwb = ThisWorkbook
    Set lltranssh = currwb.Worksheets(LLSHEET)
    Set dicttranssh = currwb.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    
End Sub

'Clear all the data in the current workbook
Private Sub CMD_ClearData_Click()
    Dim impObj As IImpSpecs
    Set impObj = ImpSpecs.Create([F_ImportRep], Me, currwb)
    impObj.ClearData
End Sub

'Clear Historic of the geobase
Private Sub CMD_ClearGeo_Click()
    Dim geoObj As ILLGeo 'Geo object

    Set geoObj = LLGeo.Create(currwb.Worksheets(GEOSHEET))

    If MsgBox(tradmess.TranslatedValue("MSG_HistoricDelete"), _
              vbExclamation + vbYesNo, _
              tradmess.TranslatedValue("MSG_DeleteHistoric")) = vbYes Then
        
        geoObj.ClearHistoric

        MsgBox tradmess.TranslatedValue("MSG_Done"), _
                vbInformation, _
                tradmess.TranslatedValue("MSG_DeleteHistoric")
    End If
End Sub

'Open the export data form for exports
Private Sub CMD_ExportData_Click()
    Me.Hide
    ClickExportMigration
End Sub

'Import historic geobase
Private Sub CMD_ImportGeoHistoric_Click()
    Dim impObj As IImpSpecs
    Set impObj = ImpSpecs.Create([F_ImportRep], Me, currwb)

    impObj.ImportGeobase histoOnly:=True
End Sub

'Leave the advanced form
Private Sub CMD_ImportMigQuit_Click()
    Me.Hide
End Sub

'Show import report
Private Sub CMD_ImportMigRep_Click()
    Dim impObj As IImpSpecs
    Set impObj = ImpSpecs.Create([F_ImportRep], Me, currwb)

    Me.Hide
    impObj.ShowReport
End Sub

Private Sub UserForm_Initialize()

    'Manage language of the userform
    Initialize

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me

    Me.width = 250
    Me.height = 450
End Sub
