Attribute VB_Name = "TestDiseaseIntegration"
Attribute VB_Description = "Integration tests covering disease add/export/import/remove workflows"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Integration tests covering disease add/export/import/remove workflows")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const ANCHOR_SHEET As String = "Variables"
Private Const DROPDOWN_SHEET As String = "__dropdowns"
Private Const TRANSLATION_SHEET As String = "IntegrationTranslations"
Private Const IMPORT_SHEET As String = "IntegrationImport"
Private Const LANGUAGES_LIST As String = "__data_languages"
Private Const STATUS_LIST As String = "__var_status"
Private Const CHOICES_LIST As String = "__lst_choices"
Private Const PROHIBITED_LIST As String = "__prohibited_diseases_list"
Private Const DISEASES_LIST As String = "__diseases_list"
Private Const VARIABLE_NAME_RANGE As String = "__Col__Variables"
Private Const MARKER_NAME_PREFIX As String = "DISSHEET"

Private Assert As CustomTest
Private Builder As DiseaseSheet
Private Importer As DiseaseImporter
Private Exporter As DiseaseExporter
Private ExportManager As DiseaseExportWorkbook
Private AppGuard As ApplicationState
Private Dropdowns As DropdownLists
Private VariablesManager As MasterSetupVariables
Private TranslationTable As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseIntegration"
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
    Set Importer = Nothing
    Set Exporter = Nothing
    Set ExportManager = Nothing
    Set AppGuard = Nothing
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

'@TestMethod("DiseaseIntegration")
Public Sub TestAddExportImportRemove()
    CustomTestSetTitles Assert, "DiseaseIntegration", "TestAddExportImportRemove"

    Dim diseaseWksh As Worksheet
    Dim diseaseTable As ListObject
    Dim exportBook As Workbook
    Dim logger As DiseaseLogger
    Dim summary As DiseaseImportSummary
    Dim entries As BetterArray
    Dim importTable As ListObject
    Dim manager As DiseaseWorksheetManager
    Dim lastCell As Range

    On Error GoTo Fail

    Set diseaseWksh = Builder.Build("Alpha")
    Set diseaseTable = diseaseWksh.ListObjects(1)

    PopulateDiseaseTable diseaseTable

    Set exportBook = Exporter.BuildDiseaseWorkbook(diseaseWksh, TranslationTable, _
                                                   "Alpha", diseaseWksh.Cells(2, 2).Value, _
                                                   HiddenNames.Create(diseaseWksh).ValueAsString("__Var_DISCODE"))

    Assert.AreEqual "Alpha", exportBook.Worksheets("Metadata").Cells(3, 2).Value, "Metadata should reference disease name"
    Assert.AreEqual "LabelA", exportBook.Worksheets("Dictionary").Cells(2, 4).Value, "Dictionary should capture existing variables"

    exportBook.Close SaveChanges:=False

    Set importTable = PrepareImportTable()
    Set logger = DiseaseLogger.Create()

    Set summary = Importer.MergeDisease(diseaseTable, importTable, True, DiseaseImportPriority_Foreign, logger)

    Assert.AreEqual "LabelAUpdated", diseaseTable.DataBodyRange.Cells(1, 4).Value, "Merge should update existing variable label"
    'The new variables land on the first free lines under var_a.
    Assert.AreEqual "var_c", diseaseTable.DataBodyRange.Cells(2, 2).Value, "Merge should land new variables on the first free line"
    Assert.AreEqual "var_d", diseaseTable.DataBodyRange.Cells(3, 2).Value, "Merge should land every new variable"
    Assert.IsTrue summary.RequiresReport, "Summary should indicate report requirement"

    Assert.IsTrue logger.HasEntries, "Logger should capture merge operations"
    Set entries = logger.Entries
    Assert.IsTrue entries.Length >= 3, "Logger should contain multiple entries for merge operations"

    'The rows the merge appended carry the dotted frame.
    Set lastCell = diseaseTable.DataBodyRange.Cells(diseaseTable.ListRows.Count, 1)
    Assert.AreEqual CLng(xlDot), CLng(lastCell.Borders(xlEdgeBottom).LineStyle), _
                    "A row appended by the merge should carry the dotted frame"

    'The rows the ribbon adds carry it too.
    MasterSetupHelpers.ManageRows diseaseWksh, True
    Set lastCell = diseaseTable.DataBodyRange.Cells(diseaseTable.ListRows.Count, 1)
    Assert.AreEqual CLng(xlDot), CLng(lastCell.Borders(xlEdgeBottom).LineStyle), _
                    "A row added through Add Rows should carry the dotted frame"
    Assert.AreEqual CLng(xlDot), CLng(lastCell.Borders(xlEdgeTop).LineStyle), _
                    "A row added through Add Rows should carry its top edge"

    Set manager = DiseaseWorksheetManager.Create()
    Assert.IsTrue manager.RemoveWorksheet(ThisWorkbook, "Alpha"), "Worksheet manager should remove disease sheet"
    Assert.IsFalse WorksheetExists("Alpha"), "Disease worksheet should be removed"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddExportImportRemove", Err.Number, Err.Description
End Sub

'@section Helpers
'===============================================================================

'The migration round trip is TestMasterSetupMigration's, over the class
'that owns the file format at both ends.

Private Sub PrepareEnvironment()
    Dim dropdownSheet As Worksheet
    Dim translationSheet As Worksheet
    Dim variablesSheet As Worksheet
    Dim variableTable As ListObject
    Dim data As Variant

    Set variablesSheet = EnsureWorksheet(ANCHOR_SHEET)
    ClearWorksheet variablesSheet

    variablesSheet.Range("A1").Value = "Variable Order"
    variablesSheet.Range("B1").Value = "Variable Section"
    variablesSheet.Range("C1").Value = "Variable Name"
    variablesSheet.Range("B2").Value = "demographics"
    variablesSheet.Range("C2").Value = "var_a"
    variablesSheet.Range("B3").Value = "symptoms"
    variablesSheet.Range("C3").Value = "var_b"

    Set variableTable = variablesSheet.ListObjects.Add(xlSrcRange, variablesSheet.Range("A1:C3"), _
                                                       XlListObjectHasHeaders:=xlYes)
    variableTable.Name = "TST_IntegrationVariables"

    Set VariablesManager = MasterSetupVariables.Create(variableTable)
    RegisterVariableName variableTable

    Set dropdownSheet = EnsureWorksheet(DROPDOWN_SHEET)
    ClearWorksheet dropdownSheet

    Set Dropdowns = DropdownLists.Create(dropdownSheet)
    AddDropdownList Dropdowns, LANGUAGES_LIST, Array("ENG", "FRA", "ESP")
    AddDropdownList Dropdowns, STATUS_LIST, Array("core", "optional", "hidden")
    AddDropdownList Dropdowns, CHOICES_LIST, Array("choice_age", "choice_fever", "choice_other", "choice_new")
    AddDropdownList Dropdowns, PROHIBITED_LIST, Array("Variables", "Translations")
    AddDropdownList Dropdowns, DISEASES_LIST, Array("", "")

    Set translationSheet = EnsureWorksheet(TRANSLATION_SHEET)
    ClearWorksheet translationSheet

    data = Array( _
        Array("tag", "ENG"), _
        Array("selectValue", "Select a value"), _
        Array("infoSelectLang", "Select language"), _
        Array("varOrder", "Variable Order"), _
        Array("varSection", "Variable Section"), _
        Array("varName", "Variable Name"), _
        Array("varLabel", "Main Label"), _
        Array("varChoice", "Choice"), _
        Array("choiceVal", "Choice Values"), _
        Array("varStatus", "Status"), _
        Array("errLang", "Please select a language") _
    )

    translationSheet.Range("A1").Resize(UBound(data) + 1, 2).Value = data
    translationSheet.ListObjects.Add SourceType:=xlSrcRange, _
                                      Source:=translationSheet.Range("A1").Resize(UBound(data) + 1, 2), _
                                      XlListObjectHasHeaders:=xlYes

    Set TranslationTable = translationSheet.ListObjects(1)

    Set Builder = DiseaseSheet.Create(ThisWorkbook, Dropdowns, VariablesManager)
    Set Importer = DiseaseImporter.Create()
    Set ExportManager = DiseaseExportWorkbook.Create()
    Set AppGuard = ApplicationState.Create(Application)
    Set Exporter = DiseaseExporter.Create(ExportManager, AppGuard)
End Sub

Private Sub CleanupEnvironment()
    DeleteWorksheetSafe "Alpha"
    DeleteWorksheetSafe DROPDOWN_SHEET
    DeleteWorksheetSafe TRANSLATION_SHEET
    DeleteWorksheetSafe IMPORT_SHEET
    DeleteWorksheetSafe "Translations"
    DeleteWorksheetSafe "Choices"
    DeleteWorksheetSafe "__dis_import"
    ClearWorksheetSafe ANCHOR_SHEET

    DeleteNameSafe VARIABLE_NAME_RANGE
    DeleteNameSafe MARKER_NAME_PREFIX & "001"
    DeleteNameSafe MARKER_NAME_PREFIX & "002"
End Sub

Private Sub PopulateDiseaseTable(ByVal table As ListObject)
    table.ListRows.Add
    table.ListRows(1).Range.Value = Array(1, "var_a", "demographics", "LabelA", "choice_age", "0-4 | 5-14", "core")
    table.ListRows(2).Range.Value = Array(2, "var_b", "symptoms", "LabelB", "choice_fever", "yes | no", "core")
End Sub

Private Function PrepareImportTable() As ListObject
    Dim importSheet As Worksheet
    Dim header As Variant
    Dim rows As Variant
    Dim tableRange As Range

    Set importSheet = EnsureWorksheet(IMPORT_SHEET)
    ClearWorksheet importSheet

    header = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    rows = RowsToMatrix(Array( _
        Array(1, "var_a", "demographics", "LabelAUpdated", "choice_age", "0-4 | 5-14", "core"), _
        Array(3, "var_c", "history", "LabelC", "choice_other", "low | high", "optional"), _
        Array(4, "var_d", "history", "LabelD", "choice_new", "alpha | beta", "core") _
    ))

    WriteMatrix importSheet.Range("A1"), header
    WriteMatrix importSheet.Range("A2"), rows

    Set tableRange = importSheet.Range("A1").Resize(UBound(rows, 1) + 1, UBound(rows, 2))
    Set PrepareImportTable = importSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                         XlListObjectHasHeaders:=xlYes)
End Function

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
    WorksheetExists = Not FindWorksheet(ThisWorkbook, sheetName) Is Nothing
End Function

Private Function FindWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    For Each sheet In targetBook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = sheet
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
