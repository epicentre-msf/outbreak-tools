Attribute VB_Name = "TestSectionBuilder"
Attribute VB_Description = "Tests for SectionBuilder class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for SectionBuilder class")

Option Explicit

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private Dict As ILLdictionary
Private Specs As LLVarContextSpecsStub
Private TargetSheet As Worksheet
Private DropStub As DropdownListsStub
Private CustDropStub As DropdownListsStub

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "SectionBuilder"
Private Const DICTIONARY_SHEET As String = "DictFixture"
Private Const VLIST_SEC_COL As Long = 2
Private Const VLIST_SUBSEC_COL As Long = 3


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestSectionBuilder"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Dim transStub As LinelistTranslationCounterStub
    Dim fmtStub As LLFormatStub
    Dim formulaStub As FormulaDataStub

    BusyApp
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare

    Set transStub = New LinelistTranslationCounterStub
    transStub.Initialise
    Set fmtStub = New LLFormatStub
    Set formulaStub = New FormulaDataStub

    Set TargetSheet = FixtureWorkbook.Worksheets.Add

    Set DropStub = New DropdownListsStub
    DropStub.Initialise TargetSheet
    Set CustDropStub = New DropdownListsStub
    CustDropStub.Initialise TargetSheet

    Set Specs = New LLVarContextSpecsStub
    Specs.SetDictionary Dict
    Specs.SetDesignFormat fmtStub
    Specs.SetTranslation transStub
    Specs.SetFormulaData formulaStub

    Assert.BeginTest
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then
            TestHelpers.DeleteWorkbook FixtureWorkbook
        End If
    On Error GoTo 0

    Set Dict = Nothing
    Set Specs = Nothing
    Set TargetSheet = Nothing
    Set DropStub = Nothing
    Set CustDropStub = Nothing
    Set FixtureWorkbook = Nothing
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("SectionBuilder")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As ISectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestCreateRaisesWithoutSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRaisesWithoutSpecs"
    On Error GoTo ExpectError

    Dim sut As ISectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Nothing, _
        wksh:=TargetSheet)

    Assert.LogFailure "Create should raise when specs is Nothing."
    Exit Sub

ExpectError:
    Assert.IsTrue Err.Number <> 0, "An error should have been raised"
    Err.Clear
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestCreateRaisesWithoutWorksheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRaisesWithoutWorksheet"
    On Error GoTo ExpectError

    Dim sut As ISectionBuilder
    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=Nothing)

    Assert.LogFailure "Create should raise when wksh is Nothing."
    Exit Sub

ExpectError:
    Assert.IsTrue Err.Number <> 0, "An error should have been raised"
    Err.Clear
End Sub


'@section Build tests
'===============================================================================

'@TestMethod("SectionBuilder")
Public Sub TestBuildVListWritesSectionName()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildVListWritesSectionName"
    On Error GoTo TestFail

    Dim sut As ISectionBuilder
    Dim startRow As Long

    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    startRow = FindSheetStartRow("vlist1D-sheet1")
    Assert.IsTrue startRow > 0, "Sheet vlist1D-sheet1 should exist in the dictionary"

    sut.Build "vlist1D-sheet1", startRow

    ' Verify that at least one section name was written to column VLIST_SEC_COL (=2)
    Dim sectionCount As Long
    sectionCount = Application.WorksheetFunction.CountIf( _
        TargetSheet.Columns(VLIST_SEC_COL), "Controls")

    Assert.IsTrue sectionCount > 0, _
                  "Section name 'Controls' should be written to column 2"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildVListWritesSectionName", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestBuildVListWritesSubSectionName()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildVListWritesSubSectionName"
    On Error GoTo TestFail

    Dim sut As ISectionBuilder
    Dim startRow As Long

    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    startRow = FindSheetStartRow("vlist1D-sheet1")
    sut.Build "vlist1D-sheet1", startRow

    ' Verify that a subsection name was written to column VLIST_SUBSEC_COL (=3)
    Dim subSectionCount As Long
    subSectionCount = Application.WorksheetFunction.CountIf( _
        TargetSheet.Columns(VLIST_SUBSEC_COL), "Date validation")

    Assert.IsTrue subSectionCount > 0, _
                  "Subsection name 'Date validation' should be written to column 3"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildVListWritesSubSectionName", Err.Number, Err.Description
End Sub


'@TestMethod("SectionBuilder")
Public Sub TestBuildVListWritesVariableLabels()
    CustomTestSetTitles Assert, TESTMODULE, "TestBuildVListWritesVariableLabels"
    On Error GoTo TestFail

    Dim sut As ISectionBuilder
    Dim startRow As Long
    Dim labelCol As Long

    Set sut = SectionBuilder.Create( _
        layer:=SectionBuilderModeVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    startRow = FindSheetStartRow("vlist1D-sheet1")
    sut.Build "vlist1D-sheet1", startRow

    ' VList labels are written at column 4 (VLIST_START_COL - 1)
    labelCol = 4
    Dim labelCount As Long
    labelCount = Application.WorksheetFunction.CountA(TargetSheet.Columns(labelCol))

    Assert.IsTrue labelCount > 0, _
                  "Variable labels should be written to column 4"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestBuildVListWritesVariableLabels", Err.Number, Err.Description
End Sub


'@section Helpers
'===============================================================================

'@description Find the first row in the dictionary DataRange where sheet name matches.
Private Function FindSheetStartRow(ByVal sheetName As String) As Long
    Dim sheetRng As Range
    Dim endRow As Long
    Dim rowIdx As Long

    Set sheetRng = Dict.DataRange("sheet name")
    endRow = Dict.Data.DataEndRow()

    For rowIdx = 1 To endRow
        If CStr(sheetRng.Cells(rowIdx, 1).Value) = sheetName Then
            FindSheetStartRow = rowIdx
            Exit Function
        End If
    Next rowIdx

    FindSheetStartRow = 0
End Function
