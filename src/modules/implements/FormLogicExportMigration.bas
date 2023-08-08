Attribute VB_Name = "FormLogicExportMigration"
Attribute VB_Description = "Form implementation of Exports for Migration"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
Option Explicit

Private tradform As ITranslation  'Translation of forms
Private tradmess As ITranslation 'Translation messages
Private geoObj As ILLGeo 'Geo object
Private expOut As IOutputSpecs
Private currwb As Workbook 'Current workbook
Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const GEOSHEET As String = "Geo"

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet
    Dim geosh As Worksheet
    Dim expsh As Worksheet

    Set currwb = ThisWorkbook
    Set lltranssh = currwb.Worksheets(LLSHEET)
    Set dicttranssh = currwb.Worksheets(TRADSHEET)
    Set geosh = currwb.Worksheets(GEOSHEET)

    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradform = lltrads.TransObject(TranslationOfForms)
    Set tradmess = lltrads.TransObject()
    Set geoObj = LLGeo.Create(geosh)
End Sub

'Export data for migration

'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor
    Application.DisplayAlerts = False
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
    Application.DisplayAlerts = True
End Sub

Private Sub CreateExport()

    Dim scope As Byte
    Dim folderPath As String
    Dim shouldQuit As Byte
    
    'Add Error management
    'On Error GoTo errHand
    
    scope = ExportAll
    BusyApp cursor:=xlNorthwestArrow
    
    InitializeTrads
    Set expOut = OutputSpecs.Create(currwb, scope)
    folderPath = expOut.ExportFolder()

    'Export for migration
    If Me.CHK_ExportMigData.Value Then expOut.Save tradmess
    'Export the geobase
    If Me.CHK_ExportMigGeo.Value Then expOut.SaveGeo geoObj:=geoObj, onlyHistoric:=False
    'Export only historic data of the geobase (onlyHistoric:=True)
    If Me.CHK_ExportMigGeoHistoric.Value Then expOut.SaveGeo geoObj:=geoObj, onlyHistoric:=True
    NotBusyApp

    'Ask the user if I should quit the form
    shouldQuit = MsgBox(tradmess.TranslatedValue("MSG_FinishedExports"), _
                        vbQuestion + vbYesNo, _
                        tradmess.TranslatedValue("MSG_Migration"))
    If shouldQuit = vbYes Then Me.Hide
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox tradmess.TranslatedValue("MSG_ErrExportData"), _
            vbOKOnly + vbCritical, _
            tradmess.TranslatedValue("MSG_Error")
    expOut.CloseAll
    On Error GoTo 0
    NotBusyApp
End Sub

Private Sub CMD_ExportMig_Click()
    CreateExport
End Sub

Private Sub CMD_ExportMigQuit_Click()
    Me.Hide
End Sub

Private Sub LBL_Previous_Click()
    Me.Hide
    F_Advanced.Show
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    'Manage language
    Me.Caption = tradform.TranslatedValue(Me.Name)

    tradform.TranslateForm Me

    Me.width = 200
    Me.height = 300
End Sub
