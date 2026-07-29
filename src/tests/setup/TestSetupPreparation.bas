Attribute VB_Name = "TestSetupPreparation"
Attribute VB_Description = "Unit tests for the SetupPreparation orchestration helper"

Option Explicit

'===============================================================================
' @ModuleDescription Unit tests for the SetupPreparation orchestration helper.
'
' @description This module validates that the SetupPreparation class correctly
'   registers dropdown list objects on the __variables worksheet, initialises
'   the updated values registry on the __updated worksheet, and applies data
'   validations to all setup tables (Dictionary, Exports, Analysis). Each test
'   creates a fresh fixture workbook with the required sheets and ListObjects,
'   calls Subject.Prepare, then asserts the expected side effects.
'
' @depends SetupPreparation, Development, BetterArray, CustomTest,
'   TestHelpersLite, DropdownLists, UpdatedValues, CustomTable, Checking
'
' The fixture workbook is rebuilt before every test via TestInitialize so each
' test runs in isolation. Dropdown content is verified by checking the
' BetterArray returned from Subject.Dropdowns.Values. Registry initialisation
' is verified by scanning ListObjects on the registry sheet for status and
' rngname columns. Validations are verified by inspecting the
' Range.Validation.Formula1 property on target columns.
'===============================================================================

'@Folder("CustomTests.Setup")
'@ModuleDescription("Validates that SetupPreparation registers dropdowns, initialises updated values, and applies setup validations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private DropdownSheet As Worksheet
Private RegistrySheet As Worksheet
Private VariablesSheet As Worksheet
Private ChoicesSheet As Worksheet
Private ExportsSheet As Worksheet
Private AnalysisSheet As Worksheet
Private CheckingSheet As Worksheet
Private Subject As SetupPreparation
Private Manager As Development
Private DevSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DROPDOWN_SHEET_NAME As String = "__variables"
Private Const UPDATED_SHEET_NAME As String = "__updated"
Private Const VARIABLES_SHEET_NAME As String = "Dictionary"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const STATUS_UPDATED As String = "yes"
Private Const TAG_WATCH_UPDATE As String = "watch for update"

'@section Lifecycle
'===============================================================================
'@description Module and test-level setup and teardown routines.

'@sub-title Initialise the test module and configure the output sheet.
'@details
'Disables screen updating via BusyApp, ensures the output sheet
'exists, creates the CustomTest assertion object, and sets the module name
'for grouped test reporting.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupPreparation"
End Sub

'@sub-title Tear down the test module after all tests have run.
'@details
'Prints accumulated results to the output sheet and restores normal
'Application state. Silently ignores errors during PrintResults to
'ensure cleanup always completes.
'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub

'@sub-title Build a fresh fixture workbook before each test.
'@details
'Creates a new workbook with all required worksheets (dropdown, registry,
'Dictionary, Choices, Exports, Analysis, checking, formatter, formula, pass,
'Dev), populates them with ListObjects matching the setup workbook layout,
'defines the required named ranges (RNG_SelectTable, RNG_CheckingFilter,
'ModulesCodes, ClassesImplementation), and creates the Subject and Manager
'instances.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    Set FixtureWorkbook = NewWorkbook
    Set DropdownSheet = EnsureWorksheet(DROPDOWN_SHEET_NAME, FixtureWorkbook)
    Set RegistrySheet = EnsureWorksheet(UPDATED_SHEET_NAME, FixtureWorkbook, True, xlSheetVeryHidden)
    Set VariablesSheet = EnsureWorksheet(VARIABLES_SHEET_NAME, FixtureWorkbook)
    Set ChoicesSheet = EnsureWorksheet(CHOICES_SHEET_NAME, FixtureWorkbook)
    Set ExportsSheet = EnsureWorksheet("Exports", FixtureWorkbook)
    Set AnalysisSheet = EnsureWorksheet("Analysis", FixtureWorkbook)
    Set CheckingSheet = EnsureWorksheet("__checkRep", FixtureWorkbook)
    EnsureWorksheet "__formatter", FixtureWorkbook
    EnsureWorksheet "__formula", FixtureWorkbook
    EnsureWorksheet "__pass", FixtureWorkbook
    EnsureWorksheet "Translations", FixtureWorkbook

    BuildWatchedTable VariablesSheet, "Tab_Dictionary", _
        Array("sheet type", "editable label", "status", "personal identifier", "variable type", "variable format", _
              "control", "register book", "unique", "alert", "lock cells"), _
        Array("vlist1D", "yes", "mandatory", "yes", "integer", "integer", _
              "choice_manual", "print, horizontal header", "yes", "error", "yes")
    BuildWatchedTable ChoicesSheet, "Tab_Choices", Array("choice"), Array("option_a"), startRow:=2, startColumn:=1
    BuildSimpleTable ExportsSheet, "Tab_Export", _
        Array("status", "file format", "password", "include personal identifiers", "include p-codes", _
              "header format", "export metadata", "export analyses sheets"), startRow:=2, startColumn:=1
    BuildAnalysisTables AnalysisSheet

    EnsureWorkbookName FixtureWorkbook, "RNG_SelectTable", AnalysisSheet.Cells(1, 1)
    EnsureWorkbookName FixtureWorkbook, "RNG_CheckingFilter", CheckingSheet.Cells(1, 1)

    Set DevSheet = EnsureWorksheet("Dev", FixtureWorkbook)
    EnsureLocalName DevSheet, "ModulesCodes", DevSheet.Cells(1, 1)
    EnsureLocalName DevSheet, "ClassesImplementation", DevSheet.Cells(2, 1)

    Set Subject = SetupPreparation.Create(FixtureWorkbook)
    Set Manager = Development.Create(DevSheet)
End Sub

'@sub-title Destroy all fixture objects and close the fixture workbook.
'@details
'Flushes pending assertions, closes and deletes the fixture workbook, and
'releases all module-level object references in reverse order of creation.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Subject = Nothing
    Set Manager = Nothing
    Set DevSheet = Nothing
    Set CheckingSheet = Nothing
    Set AnalysisSheet = Nothing
    Set ExportsSheet = Nothing
    Set ChoicesSheet = Nothing
    Set VariablesSheet = Nothing
    Set RegistrySheet = Nothing
    Set DropdownSheet = Nothing
    Set FixtureWorkbook = Nothing
End Sub

'@section Test Methods
'===============================================================================
'@description Verify the SetupPreparation preparation workflow.

'@sub-title Verify that Prepare registers standard dropdown lists.
'@details
'Calls Subject.Prepare, then retrieves the __yesno and __formats
'dropdowns via Subject.Dropdowns.Values. Asserts that __yesno contains
'exactly two entries ("yes" and "no") and that __formats includes both
'"percentage2" and "text", confirming that RegisterAllDropdowns populated
'the expected dropdown list objects.
'@TestMethod("SetupPreparation")
Public Sub TestPrepareAddsDropdowns()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareAddsDropdowns"
    On Error GoTo Fail

    Subject.Prepare Manager

    Dim yesNo As BetterArray
    Dim formats As BetterArray

    Set yesNo = Subject.Dropdowns.Values("__yesno")
    Assert.IsFalse yesNo Is Nothing, "__yesno dropdown should be created"
    Assert.AreEqual CLng(2), yesNo.Length, "__yesno dropdown should contain two entries"
    Assert.IsTrue ContainsValue(yesNo, "yes"), "__yesno dropdown should contain 'yes'"
    Assert.IsTrue ContainsValue(yesNo, "no"), "__yesno dropdown should contain 'no'"

    Set formats = Subject.Dropdowns.Values("__formats")
    Assert.IsFalse formats Is Nothing, "__formats dropdown should be created"
    Assert.IsTrue ContainsValue(formats, "percentage2"), "__formats dropdown should include percentage variants"
    Assert.IsTrue ContainsValue(formats, "text"), "__formats dropdown should include text option"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareAddsDropdowns"
End Sub

'@sub-title Verify that Prepare initialises the updated values registry.
'@details
'Calls Subject.Prepare, then scans all ListObjects on the __updated sheet.
'For each table that has "updated" and "rngname" columns, verifies that
'every status cell is set to "yes" and that every non-empty rngname cell
'has a corresponding workbook-level defined Name. Asserts that at least one
'registry table was populated.
'@TestMethod("SetupPreparation")
Public Sub TestPrepareInitialisesUpdatedValuesRegistry()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareInitialisesUpdatedValuesRegistry"
    On Error GoTo Fail

    Subject.Prepare Manager

    Dim registryCount As Long
    Dim lo As ListObject
    Dim statusColumn As Range
    Dim rangeColumn As Range
    Dim cell As Range
    Dim definedName As Name

    Set RegistrySheet = FixtureWorkbook.Worksheets(UPDATED_SHEET_NAME)

    For Each lo In RegistrySheet.ListObjects
        On Error Resume Next
            Set statusColumn = lo.ListColumns("updated").DataBodyRange
            Set rangeColumn = lo.ListColumns("rngname").DataBodyRange
        On Error GoTo 0

        If Not statusColumn Is Nothing And Not rangeColumn Is Nothing Then
            registryCount = registryCount + 1

            For Each cell In statusColumn.Cells
                Assert.AreEqual STATUS_UPDATED, NormalizeText(CStr(cell.Value)), "Registry rows should be initialised to 'yes' on listObject " & lo.Name
            Next cell

            For Each cell In rangeColumn.Cells
                If LenB(Trim$(CStr(cell.Value))) > 0 Then
                    On Error Resume Next
                        Set definedName = FixtureWorkbook.Names(CStr(cell.Value))
                    On Error GoTo 0
                    Assert.IsFalse definedName Is Nothing, "Registry should create workbook names for watched ranges"
                End If
            Next cell
        End If

        Set statusColumn = Nothing
        Set rangeColumn = Nothing
    Next

    Assert.IsTrue registryCount > 0, "Registry should be populated when tagged columns are registered"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareInitialisesUpdatedValuesRegistry"
End Sub

'@sub-title Verify that Prepare applies data validations to the Dictionary table.
'@details
'Calls Subject.Prepare, retrieves the Tab_Dictionary ListObject from the
'Dictionary worksheet, and inspects the "sheet type" column. Asserts that
'the column has a list-type data validation whose formula references the
'"__sheet_type" dropdown name.
'@TestMethod("SetupPreparation")
Public Sub TestPrepareAppliesDictionaryValidation()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareAppliesDictionaryValidation"
    On Error GoTo Fail

    Subject.Prepare Manager

    Dim lo As ListObject
    Dim targetRange As Range

    Set lo = VariablesSheet.ListObjects("Tab_Dictionary")
    Set targetRange = lo.ListColumns("sheet type").DataBodyRange

    AssertValidationContains targetRange, "__sheet_type"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareAppliesDictionaryValidation"
End Sub

'@sub-title Verify that Prepare applies data validations to Analysis tables.
'@details
'Calls Subject.Prepare, then checks two analysis-level validations: the
'RNG_SelectTable named range should have a list validation referencing
'"__switch_tables", and the Tab_TimeSeries_Analysis table's "row" column
'should reference "__time_vars". This confirms that both named-range and
'table-column validations are applied correctly.
'@TestMethod("SetupPreparation")
Public Sub TestPrepareAppliesAnalysisValidation()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareAppliesAnalysisValidation"
    On Error GoTo Fail

    Subject.Prepare Manager

    Dim selectTable As Range
    Dim tsTable As ListObject
    Dim columnRange As Range

    Set selectTable = AnalysisSheet.Range("RNG_SelectTable")
    AssertValidationContains selectTable, "__switch_tables"

    Set tsTable = AnalysisSheet.ListObjects("Tab_TimeSeries_Analysis")
    Set columnRange = tsTable.ListColumns("row").DataBodyRange
    AssertValidationContains columnRange, "__time_vars"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareAppliesAnalysisValidation"
End Sub

'@sub-title Verify that Prepare rejects a missing manager before it writes anything.
'@details
'RibbonDev.EnsureDevelopment answers Nothing on any failure and its caller
'passes that straight through, so a user can reach Prepare with no manager.
'Calls Prepare with Nothing, asserts the raised error is
'ProjectError.ObjectNotInitialized, and asserts that neither the dropdown
'sheet nor the registry sheet gained a ListObject. Before the guard moved to
'the top of Prepare, both sheets were fully built and the error came last.
'@TestMethod("SetupPreparation")
Public Sub TestPrepareRejectsMissingManagerBeforeWriting()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareRejectsMissingManagerBeforeWriting"
    On Error GoTo Fail

    Dim raisedNumber As Long

    On Error Resume Next
        Subject.Prepare Nothing
        raisedNumber = Err.Number
        Err.Clear
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), raisedNumber, _
        "Prepare should raise ObjectNotInitialized when the manager is missing"
    Assert.AreEqual CLng(0), CLng(DropdownSheet.ListObjects.Count), _
        "Prepare should register no dropdown before it tests the manager"
    Assert.AreEqual CLng(0), CLng(RegistrySheet.ListObjects.Count), _
        "Prepare should build no registry table before it tests the manager"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareRejectsMissingManagerBeforeWriting"
End Sub

'@sub-title Verify that a second Prepare writes a dropdown this class owns again.
'@details
'Runs Prepare, replaces the content of __yesno so the sheet no longer matches
'what the class declares, then runs Prepare again and asserts the two
'declared entries are back. DropdownLists.Add exits when the name is already
'registered, so before this session a prepared workbook kept its old lists
'for ever and a new format never reached it.
'@TestMethod("SetupPreparation")
Public Sub TestSecondPrepareRefreshesAnOwnedDropdown()
    CustomTestSetTitles Assert, "SetupPreparation", "TestSecondPrepareRefreshesAnOwnedDropdown"
    On Error GoTo Fail

    Dim drift As BetterArray
    Dim yesNo As BetterArray

    Subject.Prepare Manager

    Set drift = New BetterArray
    drift.LowerBound = 1
    drift.Push "maybe"
    Subject.Dropdowns.Update drift, "__yesno", removeDuplicates:=False

    Set yesNo = Subject.Dropdowns.Values("__yesno")
    Assert.IsTrue ContainsValue(yesNo, "maybe"), "the arrange step should have changed __yesno"

    Subject.Prepare Manager

    Set yesNo = Subject.Dropdowns.Values("__yesno")
    Assert.IsFalse ContainsValue(yesNo, "maybe"), "a second Prepare should write __yesno again"
    Assert.AreEqual CLng(2), yesNo.Length, "__yesno should hold its two declared entries again"
    Assert.IsTrue ContainsValue(yesNo, "yes"), "__yesno should contain 'yes' after the refresh"
    Assert.IsTrue ContainsValue(yesNo, "no"), "__yesno should contain 'no' after the refresh"
    Exit Sub

Fail:
    ReportTestFailure "TestSecondPrepareRefreshesAnOwnedDropdown"
End Sub

'@sub-title Verify that a second Prepare leaves the placeholder dropdowns alone.
'@details
'__geo_vars is one of the seven lists seeded with two empty strings here and
'filled later by EventSetup from the workbook's own content. Fills it, runs
'Prepare again, and asserts the content survived. This is the other half of
'the refresh above: writing every list again would wipe these seven on every
'press of the dev initialise button.
'@TestMethod("SetupPreparation")
Public Sub TestSecondPrepareKeepsPlaceholderDropdowns()
    CustomTestSetTitles Assert, "SetupPreparation", "TestSecondPrepareKeepsPlaceholderDropdowns"
    On Error GoTo Fail

    Dim filled As BetterArray
    Dim geoVars As BetterArray

    Subject.Prepare Manager

    Set filled = New BetterArray
    filled.LowerBound = 1
    filled.Push "adm1", "adm2"
    Subject.Dropdowns.Update filled, "__geo_vars", removeDuplicates:=False

    Subject.Prepare Manager

    Set geoVars = Subject.Dropdowns.Values("__geo_vars")
    Assert.IsTrue ContainsValue(geoVars, "adm1"), "a second Prepare should keep what EventSetup put in __geo_vars"
    Assert.IsTrue ContainsValue(geoVars, "adm2"), "a second Prepare should keep every placeholder entry"
    Exit Sub

Fail:
    ReportTestFailure "TestSecondPrepareKeepsPlaceholderDropdowns"
End Sub

'@sub-title Verify that Prepare leaves the visibility of an existing sheet alone.
'@details
'The fixture creates __updated as very hidden. ResolveOrCreateWorksheet used
'to write the visibility on every call, which forced __variables back on
'screen after every dev initialise. Asserts the registry sheet keeps the
'state the workbook gave it.
'@TestMethod("SetupPreparation")
Public Sub TestPrepareKeepsTheVisibilityOfAnExistingSheet()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareKeepsTheVisibilityOfAnExistingSheet"
    On Error GoTo Fail

    DropdownSheet.Visible = xlSheetHidden

    Subject.Prepare Manager

    Assert.AreEqual CLng(xlSheetVeryHidden), CLng(RegistrySheet.Visible), _
        "Prepare should leave the registry sheet very hidden"
    Assert.AreEqual CLng(xlSheetHidden), CLng(DropdownSheet.Visible), _
        "Prepare should leave a hidden dropdown sheet hidden"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareKeepsTheVisibilityOfAnExistingSheet"
End Sub

'@section Helpers
'===============================================================================
'@description Private utilities for fixture construction and assertion support.

'@sub-title Ensure the test output worksheet exists in ThisWorkbook.
Private Sub AssertSheetSetup()
    EnsureWorksheet TEST_OUTPUT_SHEET, ThisWorkbook, False
End Sub

'@sub-title Build a ListObject with a "watch for update" tag row above the header.
'@details
'Writes a three-row matrix (tag row, header row, data row) starting one row
'above startRow, then creates a ListObject from the header and data rows.
'The first cell of the tag row is set to TAG_WATCH_UPDATE so the updated
'values registry can discover the table during registration. When dataRow
'is omitted, generates placeholder values ("value_1", "value_2", etc.).
'@param targetSheet Worksheet. The worksheet to write the table on.
'@param tableName String. The name to assign to the new ListObject.
'@param headers Variant. Array of column header strings.
'@param dataRow Optional Variant. Array of data values for the first row. Defaults to generated placeholders.
'@param startRow Optional Long. Header row position (1-based). Defaults to 2.
'@param startColumn Optional Long. First column position (1-based). Defaults to 1.
'@return ListObject. The newly created ListObject.
Private Function BuildWatchedTable(ByVal targetSheet As Worksheet, _
                                   ByVal tableName As String, _
                                   ByVal headers As Variant, _
                                   Optional ByVal dataRow As Variant, _
                                   Optional ByVal startRow As Long = 2, _
                                   Optional ByVal startColumn As Long = 1) As ListObject

    Dim columnCount As Long
    Dim tagRow() As Variant
    Dim valuesRow() As Variant
    Dim matrix As Variant
    Dim tableRange As Range
    Dim table As ListObject
    Dim idx As Long

    columnCount = UBound(headers) - LBound(headers) + 1

    ReDim tagRow(0 To columnCount - 1)
    tagRow(0) = TAG_WATCH_UPDATE
    For idx = 1 To columnCount - 1
        tagRow(idx) = vbNullString
    Next idx

    If IsMissing(dataRow) Then
        ReDim valuesRow(0 To columnCount - 1)
        For idx = 0 To columnCount - 1
            valuesRow(idx) = "value_" & CStr(idx + 1)
        Next idx
    Else
        valuesRow = dataRow
    End If

    matrix = RowsToMatrix(Array(tagRow, headers, valuesRow))
    WriteMatrix targetSheet.Cells(startRow - 1, startColumn), matrix

    Set tableRange = targetSheet.Range(targetSheet.Cells(startRow, startColumn), _
                                       targetSheet.Cells(startRow + 1, startColumn + columnCount - 1))
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = tableName

    Set BuildWatchedTable = table
End Function

'@sub-title Build a simple header-plus-data ListObject without a tag row.
'@details
'Writes a two-row matrix (header row, data row) starting at startRow, then
'wraps it in a ListObject. Data values are generated by appending "_value"
'to each header string.
'@param targetSheet Worksheet. The worksheet to write the table on.
'@param tableName String. The name to assign to the new ListObject.
'@param headers Variant. Array of column header strings.
'@param startRow Optional Long. Header row position (1-based). Defaults to 2.
'@param startColumn Optional Long. First column position (1-based). Defaults to 1.
Private Sub BuildSimpleTable(ByVal targetSheet As Worksheet, _
                             ByVal tableName As String, _
                             ByVal headers As Variant, _
                             Optional ByVal startRow As Long = 2, _
                             Optional ByVal startColumn As Long = 1)

    Dim dataRow() As Variant
    Dim matrix As Variant
    Dim idx As Long
    Dim tableRange As Range
    Dim lo As ListObject

    ReDim dataRow(0 To UBound(headers))
    For idx = LBound(headers) To UBound(headers)
        dataRow(idx) = headers(idx) & "_value"
    Next idx

    matrix = RowsToMatrix(Array(headers, dataRow))
    WriteMatrix targetSheet.Cells(startRow, startColumn), matrix

    Set tableRange = targetSheet.Range(targetSheet.Cells(startRow, startColumn), _
                                       targetSheet.Cells(startRow + 1, startColumn + UBound(headers)))

    Set lo = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = tableName
End Sub

'@sub-title Build all eight analysis ListObjects on a single worksheet.
'@details
'Creates the full set of analysis tables (Global Summary, Univariate,
'Bivariate, Time Series, Graph Time Series, Spatial, SpatioTemporal Specs,
'SpatioTemporal) using BuildSimpleTable at staggered row offsets. The
'column headers match those expected by SetupPreparation.ApplyAnalysisValidations.
'@param analysisSheet Worksheet. The worksheet to populate with analysis tables.
Private Sub BuildAnalysisTables(ByVal analysisSheet As Worksheet)
    BuildSimpleTable analysisSheet, "Tab_Global_Summary", Array("format"), startRow:=3, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_Univariate_Analysis", _
        Array("add missing data", "format", "add percentage", "add graph", "flip coordinates", "row"), startRow:=6, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_Bivariate_Analysis", _
        Array("add missing data", "format", "add percentage", "add Graph", "flip coordinates", "row", "column"), startRow:=10, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_TimeSeries_Analysis", _
        Array("add missing data", "format", "add percentage", "add total", "row", "column"), startRow:=14, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_Graph_TimeSeries", _
        Array("plot values or percentages", "chart type", "y-axis"), startRow:=18, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_Spatial_Analysis", _
        Array("row", "column", "add missing data", "add percentage", "add graph", "flip coordinates", "format"), startRow:=21, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_SpatioTemporal_Specs", _
        Array("spatial type"), startRow:=25, startColumn:=1
    BuildSimpleTable analysisSheet, "Tab_SpatioTemporal_Analysis", _
        Array("row", "column", "format", "flip coordinates", "add graph"), startRow:=28, startColumn:=1
End Sub

'@sub-title Check whether a BetterArray contains the expected string value.
'@details
'Performs a case-insensitive, trimmed comparison against every item in the
'array. Returns True on the first match and False when no match is found
'or when items is Nothing.
'@param items BetterArray. The collection to search.
'@param expected String. The value to look for.
'@return Boolean. True when the value is found.
Private Function ContainsValue(ByVal items As BetterArray, ByVal expected As String) As Boolean
    Dim idx As Long
    Dim candidate As Variant

    If items Is Nothing Then Exit Function

    For idx = items.LowerBound To items.UpperBound
        candidate = items.Item(idx)
        If NormalizeText(CStr(candidate)) = NormalizeText(expected) Then
            ContainsValue = True
            Exit Function
        End If
    Next idx
End Function

'@sub-title Normalise a string to lowercase trimmed form for comparisons.
'@param valueText String. The text to normalise.
'@return String. The trimmed, lowercased text.
Private Function NormalizeText(ByVal valueText As String) As String
    NormalizeText = LCase$(Trim$(valueText))
End Function

'@sub-title Define or replace a workbook-level named range.
'@details
'Deletes any existing name with the given identifier, then creates a new
'workbook-scoped name pointing to the supplied anchor cell. Used to set up
'RNG_SelectTable and RNG_CheckingFilter for the fixture workbook.
'@param wb Workbook. The workbook to define the name in.
'@param nameId String. The name identifier to create.
'@param anchor Range. The cell the name should refer to.
Private Sub EnsureWorkbookName(ByVal wb As Workbook, ByVal nameId As String, ByVal anchor As Range)
    Dim refersTo As String

    refersTo = "=" & anchor.Address(True, True, xlA1, True)
    On Error Resume Next
        wb.Names(nameId).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nameId, RefersTo:=refersTo
End Sub

'@sub-title Define or replace a worksheet-level named range.
'@details
'Deletes any existing worksheet-scoped name with the given identifier,
'then creates a new one pointing to the supplied anchor cell. Used to set
'up the ModulesCodes and ClassesImplementation names on the Dev sheet.
'@param targetSheet Worksheet. The worksheet to define the name on.
'@param nameId String. The name identifier to create.
'@param anchor Range. The cell the name should refer to.
Private Sub EnsureLocalName(ByVal targetSheet As Worksheet, ByVal nameId As String, ByVal anchor As Range)
    Dim refersTo As String

    refersTo = "=" & anchor.Address(True, True, xlA1, True)
    On Error Resume Next
        targetSheet.Names(nameId).Delete
    On Error GoTo 0
    targetSheet.Names.Add Name:=nameId, RefersTo:=refersTo
End Sub

'@sub-title Assert that a range has a list validation referencing the expected dropdown tag.
'@details
'Verifies that the target range is not Nothing, that its validation type is
'xlValidateList, and that its Formula1 contains the expected dropdown name
'(case-insensitive). Used by test methods to confirm that
'SetupPreparation wired the correct dropdown to each table column.
'@param targetRange Range. The range whose validation to inspect.
'@param expectedTag String. The dropdown name expected in the validation formula.
Private Sub AssertValidationContains(ByVal targetRange As Range, ByVal expectedTag As String)
    Dim validationFormula As String

    Assert.IsFalse targetRange Is Nothing, "Validation target range should exist"

    With targetRange.Validation
        Assert.AreEqual xlValidateList, .Type, "Validation should be a list"
        validationFormula = NormalizeText(CStr(.Formula1))
        Assert.IsTrue InStr(1, validationFormula, NormalizeText(expectedTag), vbTextCompare) > 0, _
            "Validation formula should reference dropdown '" & expectedTag & "'"
    End With
End Sub

'@sub-title Log a test failure with error details and clear the error state.
'@details
'Formats a failure message including the error number, source, and
'description, logs it through the Assert object, then clears the error.
'Used as the common error handler across all test methods.
'@param context String. The name of the failing test method.
Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
