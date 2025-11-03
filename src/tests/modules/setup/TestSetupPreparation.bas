Attribute VB_Name = "TestSetupPreparation"
Attribute VB_Description = "Unit tests for the SetupPreparation orchestration helper"

Option Explicit

'@Folder("CustomTests.Setup")
'@ModuleDescription("Validates that SetupPreparation registers dropdowns, initialises updated values, and applies setup validations")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, ProcedureNotUsed

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private DropdownSheet As Worksheet
Private RegistrySheet As Worksheet
Private VariablesSheet As Worksheet
Private ChoicesSheet As Worksheet
Private ExportsSheet As Worksheet
Private AnalysisSheet As Worksheet
Private CheckingSheet As Worksheet
Private Subject As ISetupPreparation
Private Manager As IDevelopment
Private DevSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DROPDOWN_SHEET_NAME As String = "__variables"
Private Const UPDATED_SHEET_NAME As String = "__updated"
Private Const VARIABLES_SHEET_NAME As String = "Dictionary"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const STATUS_DEFAULT As String = "no"
Private Const STATUS_UPDATED As String = "yes"
Private Const TAG_WATCH_UPDATE As String = "watch for update"

'@ModuleInitialize
Public Sub ModuleInitialize()
    TestHelpers.BusyApp
    AssertSheetSetup
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestSetupPreparation"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    TestHelpers.RestoreApp
End Sub

'@TestInitialize
Public Sub TestInitialize()
    TestHelpers.BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set DropdownSheet = TestHelpers.EnsureWorksheet(DROPDOWN_SHEET_NAME, FixtureWorkbook)
    Set RegistrySheet = TestHelpers.EnsureWorksheet(UPDATED_SHEET_NAME, FixtureWorkbook, True, xlSheetVeryHidden)
    Set VariablesSheet = TestHelpers.EnsureWorksheet(VARIABLES_SHEET_NAME, FixtureWorkbook)
    Set ChoicesSheet = TestHelpers.EnsureWorksheet(CHOICES_SHEET_NAME, FixtureWorkbook)
    Set ExportsSheet = TestHelpers.EnsureWorksheet("Exports", FixtureWorkbook)
    Set AnalysisSheet = TestHelpers.EnsureWorksheet("Analysis", FixtureWorkbook)
    Set CheckingSheet = TestHelpers.EnsureWorksheet("__checkRep", FixtureWorkbook)
    TestHelpers.EnsureWorksheet "__formatter", FixtureWorkbook
    TestHelpers.EnsureWorksheet "__formula", FixtureWorkbook
    TestHelpers.EnsureWorksheet "__pass", FixtureWorkbook

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

    Set DevSheet = TestHelpers.EnsureWorksheet("Dev", FixtureWorkbook)
    EnsureLocalName DevSheet, "ModulesCodes", DevSheet.Cells(1, 1)
    EnsureLocalName DevSheet, "ClassesImplementation", DevSheet.Cells(2, 1)

    Set Subject = SetupPreparation.Create(FixtureWorkbook)
    Set Manager = Development.Create(DevSheet)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        TestHelpers.DeleteWorkbook FixtureWorkbook
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
        If lo.Name = "__UpLo__Names__" Then GoTo NextLo
        On Error Resume Next
            Set statusColumn = lo.ListColumns("updated").DataBodyRange
            Set rangeColumn = lo.ListColumns("rngname").DataBodyRange
        On Error GoTo 0

        If Not statusColumn Is Nothing And Not rangeColumn Is Nothing Then
            registryCount = registryCount + 1

            For Each cell In statusColumn.Cells
                Assert.AreEqual STATUS_UPDATED, NormalizeText(CStr(cell.Value)), "Registry rows should be initialised to 'no' on listObject " & lo.Name
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
    NextLo:
    Next

    Assert.IsTrue registryCount > 0, "Registry should be populated when tagged columns are registered"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareInitialisesUpdatedValuesRegistry"
End Sub

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

'@TestMethod("SetupPreparation")
Public Sub TestPrepareAppliesAnalysisValidation()
    CustomTestSetTitles Assert, "SetupPreparation", "TestPrepareAppliesAnalysisValidation"
    On Error GoTo Fail

    Subject.Prepare Manager

    Dim selectTable As Range
    Dim tsTable As ListObject
    Dim columnRange As Range

    Set selectTable = AnalysisSheet.Range("RNG_SelectTable")
    AssertValidationContains selectTable, "__swicth_tables"

    Set tsTable = AnalysisSheet.ListObjects("Tab_TimeSeries_Analysis")
    Set columnRange = tsTable.ListColumns("row").DataBodyRange
    AssertValidationContains columnRange, "__time_vars"
    Exit Sub

Fail:
    ReportTestFailure "TestPrepareAppliesAnalysisValidation"
End Sub

'@section Helpers
'===============================================================================

Private Sub AssertSheetSetup()
    TestHelpers.EnsureWorksheet TEST_OUTPUT_SHEET, ThisWorkbook, False
End Sub

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

    matrix = TestHelpers.RowsToMatrix(Array(tagRow, headers, valuesRow))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow - 1, startColumn), matrix

    Set tableRange = targetSheet.Range(targetSheet.Cells(startRow, startColumn), _
                                       targetSheet.Cells(startRow + 1, startColumn + columnCount - 1))
    Set table = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    table.Name = tableName

    Set BuildWatchedTable = table
End Function

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

    matrix = TestHelpers.RowsToMatrix(Array(headers, dataRow))
    TestHelpers.WriteMatrix targetSheet.Cells(startRow, startColumn), matrix

    Set tableRange = targetSheet.Range(targetSheet.Cells(startRow, startColumn), _
                                       targetSheet.Cells(startRow + 1, startColumn + UBound(headers)))

    Set lo = targetSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    lo.Name = tableName
End Sub

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

Private Function NormalizeText(ByVal valueText As String) As String
    NormalizeText = LCase$(Trim$(valueText))
End Function

Private Sub EnsureWorkbookName(ByVal wb As Workbook, ByVal nameId As String, ByVal anchor As Range)
    Dim refersTo As String

    refersTo = "=" & anchor.Address(True, True, xlA1, True)
    On Error Resume Next
        wb.Names(nameId).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nameId, RefersTo:=refersTo
End Sub

Private Sub EnsureLocalName(ByVal targetSheet As Worksheet, ByVal nameId As String, ByVal anchor As Range)
    Dim refersTo As String

    refersTo = "=" & anchor.Address(True, True, xlA1, True)
    On Error Resume Next
        targetSheet.Names(nameId).Delete
    On Error GoTo 0
    targetSheet.Names.Add Name:=nameId, RefersTo:=refersTo
End Sub

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

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub
