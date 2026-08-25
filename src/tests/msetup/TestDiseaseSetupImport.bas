Attribute VB_Name = "TestDiseaseSetupImport"
Attribute VB_Description = "Tests proving an exported disease imports cleanly into a setup"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests proving an exported disease imports cleanly into a setup")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const PASS_SHEET As String = "TST_DisImp_Pass"
Private Const DICTIONARY_SHEET As String = "Dictionary"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const DISEASE_FIXTURE_SHEET As String = "DisImpAlpha"
Private Const CHOICES_TABLE As String = "Tab_Choices"
Private Const TRANSLATIONS_TABLE As String = "Tab_Translations"

Private Assert As CustomTest
Private ExportedPath As String

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseSetupImport"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        CleanupEnvironment
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseSetupImport")
Public Sub TestExportedDiseaseImportsIntoSetup()
    CustomTestSetTitles Assert, "DiseaseSetupImport", "TestExportedDiseaseImportsIntoSetup"

    Dim diseaseWksh As Worksheet
    Dim translationTable As ListObject
    Dim exporter As DiseaseExporter
    Dim pass As Passwords
    Dim service As SetupImport
    Dim sheetsList As BetterArray

    On Error GoTo Fail

    'The host setup sheets, from the same fixtures the setup import suite uses.
    PreparePasswordsFixture PASS_SHEET
    PrepareSetupDictionarySheet DICTIONARY_SHEET, "existing_var", "sheet1", 5, 1
    PrepareSetupChoicesSheet CHOICES_SHEET, 4, 1, , CHOICES_TABLE
    PrepareSetupTranslationsSheet TRANSLATIONS_SHEET, TRANSLATIONS_TABLE, "hello", "Hello", "greet", 1, 2

    'The source file: one disease, exported the way the ribbon does it.
    Set diseaseWksh = PrepareDiseaseFixture()
    Set translationTable = ThisWorkbook.Worksheets(TRANSLATIONS_SHEET).ListObjects(TRANSLATIONS_TABLE)

    Set exporter = DiseaseExporter.Create(DiseaseExportWorkbook.Create(), _
                                          ApplicationState.Create(Application))
    ExportedPath = exporter.ExportDisease(EnsureTempFolder(), diseaseWksh, translationTable, _
                                          diseaseWksh.Name, "ENG", "DISCODE1")

    Assert.IsTrue LenB(Dir(ExportedPath)) > 0, "The disease export should land on disk"

    'The import, the way a setup runs it.
    Set pass = Passwords.Create(ThisWorkbook.Worksheets(PASS_SHEET))
    pass.DisplayPrompts = False

    Set sheetsList = New BetterArray
    sheetsList.Push DICTIONARY_SHEET, CHOICES_SHEET, TRANSLATIONS_SHEET

    Set service = SetupImport.Create(ExportedPath)
    service.Import pass, sheetsList

    Assert.IsTrue ColumnHoldsValue(DICTIONARY_SHEET, "age"), _
                  "The disease variables should land in the setup dictionary"
    Assert.IsTrue ColumnHoldsValue(CHOICES_SHEET, "choice_age"), _
                  "The disease choices should land in the setup choices sheet"
    Assert.IsTrue ColumnHoldsValue(CHOICES_SHEET, "0-4"), _
                  "The choice values should land beside their list"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExportedDiseaseImportsIntoSetup", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Function PrepareDiseaseFixture() As Worksheet
    Dim sheet As Worksheet
    Dim store As HiddenNames
    Dim header As Variant
    Dim dataRows As Variant
    Dim listRange As Range

    Set sheet = EnsureWorksheet(DISEASE_FIXTURE_SHEET)
    ClearWorksheet sheet

    sheet.Cells(2, 2).Value = "ENG"

    Set store = HiddenNames.Create(sheet)
    store.EnsureName "sheetTag", "disease", HiddenNameTypeString
    store.EnsureName "__Var_DISLANG", "ENG", HiddenNameTypeString
    store.EnsureName "__Var_DISCODE", "DISCODE1", HiddenNameTypeString

    header = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    dataRows = RowsToMatrix(Array( _
        Array(1, "age", "demographics", "Age", "choice_age", "0-4 | 5-14", "core"), _
        Array(2, "fever", "symptoms", "Fever", "choice_fever", "yes | no", "core") _
    ))

    WriteMatrix sheet.Range("B4"), header
    WriteMatrix sheet.Range("B5"), dataRows

    Set listRange = sheet.Range("B4").Resize(3, 7)
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set PrepareDiseaseFixture = sheet
End Function

Private Function EnsureTempFolder() As String
    EnsureTempFolder = ThisWorkbook.Path & Application.PathSeparator & "temp"

    ' MkDir raises when the folder already exists; existing is the good case.
    On Error Resume Next
        MkDir EnsureTempFolder
    On Error GoTo 0
End Function

'@description Answer whether any used cell of the sheet carries the value.
Private Function ColumnHoldsValue(ByVal sheetName As String, ByVal expected As String) As Boolean
    Dim usedCells As Range
    Dim hit As Range

    Set usedCells = ThisWorkbook.Worksheets(sheetName).UsedRange
    Set hit = usedCells.Find(What:=expected, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    ColumnHoldsValue = Not hit Is Nothing
End Function

Private Sub CleanupEnvironment()
    On Error Resume Next
        ThisWorkbook.Worksheets(DICTIONARY_SHEET).Delete
        ThisWorkbook.Worksheets(CHOICES_SHEET).Delete
        ThisWorkbook.Worksheets(TRANSLATIONS_SHEET).Delete
        ThisWorkbook.Worksheets(DISEASE_FIXTURE_SHEET).Delete
        ThisWorkbook.Worksheets(PASS_SHEET).Delete
        ThisWorkbook.Names("RNG_PublicKey").Delete
        ThisWorkbook.Names("RNG_PrivateKey").Delete
        ThisWorkbook.Names("RNG_DebuggingPassword").Delete
        ThisWorkbook.Names("RNG_DebugMode").Delete
        ThisWorkbook.Names("RNG_Version").Delete
        ThisWorkbook.Names("RNG_LabPublicKey").Delete
        ThisWorkbook.Names("RNG_LabPrivateKey").Delete
        ThisWorkbook.Names("Passwords_ProtectedSheets").Delete
        If LenB(ExportedPath) > 0 Then
            If LenB(Dir(ExportedPath)) > 0 Then Kill ExportedPath
        End If
        ExportedPath = vbNullString
    On Error GoTo 0
End Sub
