Attribute VB_Name = "TestMasterSetupExports"
Attribute VB_Description = "Tests driving the seams of the exports module that need no picker"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests driving the seams of the exports module that need no picker")

'Three seams of MasterSetupExports and CustomMasterSetupFunctions open
'without a file picker: ImportSetupWorkbook over an open export,
'LoggerSummary over a filled DiseaseLogger, and the three worksheet functions
'over the fixture sheets. The fixture is the one TestMasterSetupImportService
'builds: the Variables table, the Choices block, the translations table with
'its helper column, and the dropdowns the builder reads.

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const VARIABLES_SHEET As String = "Variables"
Private Const CHOICES_SHEET As String = "Choices"
Private Const TRANSLATIONS_SHEET As String = "Translations"
Private Const DROPDOWN_SHEET As String = "__dropdowns"
Private Const STAGING_SHEET As String = "__dis_import"
Private Const DISEASE_NAME As String = "ExpAlpha"
Private Const LANGUAGES_LIST As String = "__data_languages"
Private Const STATUS_LIST As String = "__var_status"
Private Const CHOICES_LIST As String = "__lst_choices"
Private Const PROHIBITED_LIST As String = "__prohibited_diseases_list"
Private Const DISEASES_LIST As String = "__diseases_list"
Private Const VARIABLE_NAME_RANGE As String = "__Col__Variables"
Private Const MARKER_NAME_PREFIX As String = "DISSHEET"
Private Const CHOICES_HEADER_ROW As Long = 4

Private Assert As CustomTest
Private Builder As DiseaseSheet
Private Exporter As DiseaseExporter
Private ExportManager As DiseaseExportWorkbook
Private Dropdowns As DropdownLists
Private VariablesManager As MasterSetupVariables
Private TranslationTable As ListObject
Private ExportBook As Workbook

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestMasterSetupExports"
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
    Set Builder = Nothing
    Set Exporter = Nothing
    Set ExportManager = Nothing
    Set Dropdowns = Nothing
    Set VariablesManager = Nothing
    Set TranslationTable = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    CleanupEnvironment
    PrepareEnvironment
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
    CleanupEnvironment
End Sub

'@section Tests
'===============================================================================

'@TestMethod("MasterSetupExports")
Public Sub TestImportSetupWorkbookRebuildsTheDisease()
    CustomTestSetTitles Assert, "MasterSetupExports", "TestImportSetupWorkbookRebuildsTheDisease"

    Dim diseaseWksh As Worksheet
    Dim manager As DiseaseWorksheetManager
    Dim logger As DiseaseLogger
    Dim service As MasterSetupImportService
    Dim rebuiltTable As ListObject

    On Error GoTo Fail

    Set diseaseWksh = Builder.Build(DISEASE_NAME)
    FillDiseaseTable diseaseWksh.ListObjects(1), Array( _
        Array(1, "var_a", "demographics", "Age", "choice_age", "0 to 4 | 5 to 14", "core"), _
        Array(2, "var_new", "history", "New label", "choice_new", "low | high", "optional") _
    )
    Set ExportBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, TranslationTable, DISEASE_NAME, "ENG", "DISSHEET001")

    'The disease goes away, then comes back off the exported file through
    'the door the ribbon opens, over the managers of this workbook.
    Set manager = DiseaseWorksheetManager.Create()
    manager.RemoveWorksheet ThisWorkbook, DISEASE_NAME
    Set logger = DiseaseLogger.Create()

    Set service = MasterSetupExports.ImportSetupWorkbook(ExportBook, ThisWorkbook, logger)

    Assert.IsFalse service Is Nothing, "The door answers the service that ran the import"
    Assert.AreEqual DISEASE_NAME, service.DiseaseName, "The disease name comes from the file tag"
    Assert.IsTrue WorksheetExists(DISEASE_NAME), "The disease worksheet should be rebuilt"

    Set rebuiltTable = ThisWorkbook.Worksheets(DISEASE_NAME).ListObjects(1)
    Assert.AreEqual 2, rebuiltTable.ListRows.Count, "Both variables of the file should land"
    Assert.AreEqual "var_new", rebuiltTable.DataBodyRange.Cells(2, 2).Value, "The second variable lands on the second line"

    Assert.AreEqual 1, service.AddedVariables.Length, "The Variables table takes the variable it lacked"
    Assert.AreEqual "var_new", service.AddedVariables.Item(1), "The added variable is named"
    Assert.AreEqual 1, service.AddedChoices.Length, "The Choices sheet takes the list it lacked"
    Assert.AreEqual "choice_new", service.AddedChoices.Item(1), "The added list is named"
    Assert.IsFalse WorksheetExists(STAGING_SHEET), "The staging sheet goes away"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestImportSetupWorkbookRebuildsTheDisease", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupExports")
Public Sub TestLoggerSummaryCountsWarningsAndErrors()
    CustomTestSetTitles Assert, "MasterSetupExports", "TestLoggerSummaryCountsWarningsAndErrors"

    Dim logger As DiseaseLogger

    On Error GoTo Fail

    Assert.AreEqual vbNullString, MasterSetupExports.LoggerSummary(Nothing), "No logger answers an empty summary"

    Set logger = DiseaseLogger.Create()
    Assert.AreEqual vbNullString, MasterSetupExports.LoggerSummary(logger), "An empty logger answers an empty summary"

    logger.Record "Export", DiseaseLogInfo, "one info line"
    Assert.AreEqual vbNullString, MasterSetupExports.LoggerSummary(logger), "Info lines alone keep the summary empty"

    logger.Record "Export", DiseaseLogWarning, "first warning"
    logger.Record "Export", DiseaseLogWarning, "second warning"
    logger.Record "Export", DiseaseLogError, "one error"

    Assert.AreEqual ", 2 warning(s), 1 error(s)", MasterSetupExports.LoggerSummary(logger), _
                    "The summary counts the warnings and the errors"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLoggerSummaryCountsWarningsAndErrors", Err.Number, Err.Description
End Sub

'@TestMethod("MasterSetupExports")
Public Sub TestWorksheetFunctionsReadTheFixtureSheets()
    CustomTestSetTitles Assert, "MasterSetupExports", "TestWorksheetFunctionsReadTheFixtureSheets"

    On Error GoTo Fail

    'The functions cache the managers; the fixture was just rebuilt.
    ResetMasterSetupFunctionCaches

    Assert.AreEqual "Age", MainLabelValue("var_a", "ENG"), "The label of a variable comes from the Variables sheet"
    Assert.AreEqual "Age FR", MainLabelValue("var_a", "FRA"), "The label follows the language column of the translations"
    Assert.AreEqual vbNullString, MainLabelValue("var_unknown", "ENG"), "An unknown variable answers an empty label"
    Assert.AreEqual vbNullString, MainLabelValue("", "ENG"), "An empty name answers an empty label"

    Assert.AreEqual "demographics", VariableSectionValue("var_a", "ENG"), "The section of a variable comes from the Variables sheet"
    Assert.AreEqual "Demographie", VariableSectionValue("var_a", "FRA"), "The section follows the language column of the translations"
    Assert.AreEqual "symptoms", VariableSectionValue("var_b", "FRA"), "A section the translations table does not carry reads as it is typed"
    Assert.AreEqual vbNullString, VariableSectionValue("var_unknown", "ENG"), "An unknown variable answers an empty section"
    Assert.AreEqual vbNullString, VariableSectionValue("", "ENG"), "An empty name answers an empty section"

    Assert.AreEqual "0 to 4 | 5 to 14", ChoiceValues("choice_age", "ENG"), "The values of a choice are joined with a pipe"
    Assert.AreEqual vbNullString, ChoiceValues("choice_unknown", "ENG"), "An unknown choice answers empty values"
    Assert.AreEqual vbNullString, ChoiceValues("", "ENG"), "An empty choice answers empty values"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestWorksheetFunctionsReadTheFixtureSheets", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareEnvironment()
    Dim variablesSheet As Worksheet
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim choicesSheet As Worksheet
    Dim variableTable As ListObject
    Dim data As Variant

    'Variables sheet: the master table, in the eight columns the master carries.
    Set variablesSheet = EnsureWorksheet(VARIABLES_SHEET)
    ClearWorksheet variablesSheet
    data = RowsToMatrix(Array( _
        Array("Variable Order", "Variable Section", "Variable Name", "Label", "Default Choice", "Choices Values", "Default Status", "Comments"), _
        Array(1, "demographics", "var_a", "Age", "choice_age", "0 to 4 | 5 to 14", "core", ""), _
        Array(2, "symptoms", "var_b", "Fever", "choice_fever", "yes | no", "core", "") _
    ))
    WriteMatrix variablesSheet.Range("A1"), data
    Set variableTable = variablesSheet.ListObjects.Add(xlSrcRange, variablesSheet.Range("A1").Resize(3, 8), _
                                                       XlListObjectHasHeaders:=xlYes)
    variableTable.Name = "TST_ExportsVariables"

    Set VariablesManager = MasterSetupVariables.Create(variableTable)
    RegisterVariableName variableTable

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    Set Dropdowns = DropdownLists.Create(dropdownSheet)
    AddDropdownList Dropdowns, LANGUAGES_LIST, Array("ENG", "FRA")
    AddDropdownList Dropdowns, STATUS_LIST, Array("core", "optional", "hidden")
    AddDropdownList Dropdowns, CHOICES_LIST, Array("choice_age", "choice_fever")
    AddDropdownList Dropdowns, PROHIBITED_LIST, Array("Variables", "Translations", "Choices")
    AddDropdownList Dropdowns, DISEASES_LIST, Array("", "")

    'Translations sheet: the table sits one column right of its helper tag
    'column, the way SetupTranslationsTable wants it; the headers are the
    'languages.
    Set translationSheet = EnsureWorksheet(TRANSLATIONS_SHEET)
    ClearWorksheet translationSheet
    translationSheet.Range("A1").Value = "__TagInternal__"
    data = RowsToMatrix(Array(Array("ENG", "FRA"), Array("Age", "Age FR"), Array("Fever", "Fievre"), _
                              Array("demographics", "Demographie")))
    WriteMatrix translationSheet.Range("B1"), data
    translationSheet.ListObjects.Add SourceType:=xlSrcRange, Source:=translationSheet.Range("B1").Resize(4, 2), _
                                      XlListObjectHasHeaders:=xlYes
    Set TranslationTable = translationSheet.ListObjects(1)

    'Choices sheet: headers on row 4, the way the master file carries them.
    'The values are picked so that Excel stores none of them as a date.
    Set choicesSheet = EnsureWorksheet(CHOICES_SHEET)
    ClearWorksheet choicesSheet
    data = RowsToMatrix(Array( _
        Array("List Name", "Ordering list", "Translated Label", "Label", "Short Label"), _
        Array("choice_age", 1, "", "0 to 4", "0 to 4"), _
        Array("choice_age", 2, "", "5 to 14", "5 to 14"), _
        Array("choice_fever", 1, "", "yes", "yes"), _
        Array("choice_fever", 2, "", "no", "no") _
    ))
    WriteMatrix choicesSheet.Cells(CHOICES_HEADER_ROW, 1), data
    choicesSheet.ListObjects.Add SourceType:=xlSrcRange, _
                                 Source:=choicesSheet.Cells(CHOICES_HEADER_ROW, 1).Resize(5, 5), _
                                 XlListObjectHasHeaders:=xlYes

    Set Builder = DiseaseSheet.Create(ThisWorkbook, Dropdowns, VariablesManager)
    Set ExportManager = DiseaseExportWorkbook.Create()
    Set Exporter = DiseaseExporter.Create(ExportManager, ApplicationState.Create(Application))
End Sub

Private Sub CleanupEnvironment()
    On Error Resume Next
        If Not ExportBook Is Nothing Then ExportBook.Close SaveChanges:=False
    On Error GoTo 0
    Set ExportBook = Nothing

    ResetMasterSetupFunctionCaches

    DeleteWorksheetSafe DISEASE_NAME
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATIONS_SHEET
    DeleteWorksheetSafe CHOICES_SHEET
    DeleteWorksheetSafe STAGING_SHEET
    ClearWorksheetSafe VARIABLES_SHEET

    DeleteNameSafe VARIABLE_NAME_RANGE
    DeleteNameSafe MARKER_NAME_PREFIX & "001"
    DeleteNameSafe MARKER_NAME_PREFIX & "002"
End Sub

'@description Write the lines of a disease table over the line the builder left.
Private Sub FillDiseaseTable(ByVal table As ListObject, ByVal rows As Variant)
    Dim rowIndex As Long

    For rowIndex = LBound(rows) To UBound(rows)
        If table.ListRows.Count < rowIndex - LBound(rows) + 1 Then table.ListRows.Add
        table.ListRows(rowIndex - LBound(rows) + 1).Range.Value = rows(rowIndex)
    Next rowIndex
End Sub

Private Sub DeleteWorksheetSafe(ByVal sheetName As String)
    On Error Resume Next
        ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
End Sub

Private Sub ClearWorksheetSafe(ByVal sheetName As String)
    Dim sh As Worksheet

    On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Not sh Is Nothing Then
        ClearWorksheet sh
    End If
End Sub

Private Sub DeleteNameSafe(ByVal nameValue As String)
    On Error Resume Next
        ThisWorkbook.Names(nameValue).Delete
    On Error GoTo 0
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In ThisWorkbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            WorksheetExists = True
            Exit Function
        End If
    Next sheet
End Function

Private Sub RegisterVariableName(ByVal lo As ListObject)
    Dim store As HiddenNames

    Set store = HiddenNames.Create(ThisWorkbook)
    store.SetListObjectHeader VARIABLE_NAME_RANGE, lo, "Variable Name"
End Sub

Private Sub AddDropdownList(ByVal target As DropdownLists, ByVal listName As String, ByVal values As Variant)
    Dim listValues As BetterArray

    Set listValues = BuildBetterArray(values)
    If listValues Is Nothing Then Exit Sub

    target.Add listValues, listName
End Sub

Private Function BuildBetterArray(ByVal values As Variant) As BetterArray
    Dim arr As BetterArray
    Dim idx As Long

    If Not IsArray(values) Then Exit Function

    Set arr = New BetterArray
    arr.LowerBound = 1
    For idx = LBound(values) To UBound(values)
        arr.Push CStr(values(idx))
    Next idx

    Set BuildBetterArray = arr
End Function
