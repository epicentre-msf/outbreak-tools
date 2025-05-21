Attribute VB_Name = "Exports"
Attribute VB_Description = "Import and Export the disease to the outside world"

Option Explicit

Private Const TRADTABLE As String = "TabTransId"
Private Const TRADTABLESHEET As String = "__ribbonTranslation"
Private Const DROPSHEET As String = "__dropdowns"
Private Const RNG_FileLang As String = "RNG_FileLang"
Private Const TRANSLATIONSHEET As String = "Translations"
Private Const VARIABLESHEET As String = "Variables"
Private Const CHOICESHEET As String = "Choices"

Private trads As ITranslationObject
Private wb As Workbook

'@Folder("Import And Exports")
'@IgnoreModule UseMeaningfulName, HungarianNotation, EmptyMethod
'@ModuleDescription("Import and Export the disease to the outside world")

Private Sub InitializeTrads()

    Dim Lo As ListObject
    Dim tradTagsh As Worksheet
    Dim fileLang As String

    Set wb = ThisWorkbook
    Set tradTagsh = wb.Worksheets(TRADTABLESHEET)
    Set Lo = tradTagsh.ListObjects(TRADTABLE)
    fileLang = wb.Worksheets(TRADTABLESHEET).Range(RNG_FileLang).Value
    Set trads = Translation.Create(Lo, fileLang)
End Sub

Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.CalculateBeforeSave = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.EnableAnimations = True
    Application.Calculation = xlCalculationAutomatic
End Sub


'@EntryPoint
'@Description("Export one disease file")
Public Sub ExportToSetup()
    Attribute ExportToSetup.VB_Description="Export one disease file"

    Dim disObj As IDisease
    Dim sh As Worksheet
    Dim dropObj As IDropdownLists
    Dim outwb As Workbook

    InitializeTrads
    Set sh = ActiveSheet

    If sh.Cells(2, 4).Value <> "DISSHEET" Then
        MsgBox trads.TranslatedValue("errDisNotFound"), vbCritical, trads.TranslatedValue("error")
        Exit Sub
    End If

    On Error GoTo ExitExport

    Set dropObj = DropdownLists.Create(wb.Worksheets(DROPSHEET))
    Set disObj = Disease.Create(wb, dropObj)

    'Export one disease to setup
    disObj.ExportDisease sh.Name

    NotBusyApp
    Exit Sub

ExitExport:
    On Error Resume Next
    Set outwb = disObj.OutputWkb()
    If Not (outwb Is Nothing) Then outwb.Close saveChanges:=False
    NotBusyApp
    On Error GoTo 0
End Sub

'@EntryPoint
'@Description("Export the disease file to a flat file for migration")
Public Sub ExportForMigration()
    Attribute ExportForMigration.VB_Description="Export the disease file to a flat file for migration"

    Dim dropObj As IDropdownLists
    Dim disObj As IDisease
    Dim outwb As Workbook

    BusyApp
    'On Error GoTo ExitExport

    InitializeTrads

    Set dropObj = DropdownLists.Create(wb.Worksheets(DROPSHEET))
    Set disObj = Disease.Create(wb, dropObj)
    disObj.ExportForMigration
    NotBusyApp
    Exit Sub

ExitExport:
    On Error Resume Next
    Set outwb = disObj.OutputWkb()
    If Not (outwb Is Nothing) Then outwb.Close saveChanges:=False
    NotBusyApp
    On Error GoTo 0
End Sub



'@EntryPoint
'@Description("Import previously exported file into the disease workbook")
Public Sub ImportFlatFile()
    Attribute ImportFlatFile.VB_Description="Import previously exported file into the disease workbook"
    
    Dim disObj As IDisease
    Dim disName As String
    Dim counter As Long
    Dim lstSheets As BetterArray
    Dim io As IOSFiles






End Sub 