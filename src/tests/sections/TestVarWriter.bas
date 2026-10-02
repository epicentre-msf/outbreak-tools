Attribute VB_Name = "TestVarWriter"
Attribute VB_Description = "Tests for VarWriter class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for VarWriter class")

Option Explicit

'@description
'Drives VarWriter, which puts one dictionary variable on a data entry sheet:
'label, number format, formula, dropdown, validation, conditional formatting and
'the CRF companion row.
'
'ONE FIXTURE WORKBOOK FOR THE WHOLE MODULE
'-------------------------------------------------------------------------------
'A workbook per test meant a dictionary preparation, a choices sheet, a format
'sheet and a formula sheet per test, and that alone put the runner past its cap
'with nothing else registered. The workbook is built once and every test writes
'its own variables into the shared sheets. Each test uses variables no other test
'writes, so the sheets never need clearing.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and stops
'the whole run: no results file, and Excel left holding the staging copy. The
'setup captures its error instead and each test reports it as its own failure.
'This is the shape TestCodeTransfer uses.
'@depends VarWriter, LinelistSpecs, LLdictionary, DropdownLists, CustomTest

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private Dict As LLdictionary
Private Specs As LinelistSpecs
Private TargetSheet As Worksheet
Private HListSheet As Worksheet
Private PrintSheet As Worksheet
Private CleanPrintSheet As Worksheet
Private DropStub As DropdownLists
Private CustDropStub As DropdownLists
Private SetupError As Long
Private SetupMessage As String

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "VarWriter"
Private Const DICTIONARY_SHEET As String = "DictFixture"
Private Const CHOICES_SHEET As String = "ChoicesFixture"

'The two data entry sheets carry the names the dictionary gives them, the same
'way LLDataEntry names the worksheets it creates. LLSheets.VariableAddress drops
'the sheet prefix when the name matches, so a conditional formatting formula
'built here points where it points in a real linelist.
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const HLIST_SHEET As String = "hlist2D-sheet1"

'The date-bound row: a Min cell holding a real Date under a date number format.
Private Const DATE_BOUND_VAR As String = "date_min_v1"
Private Const DATE_BOUND_YEAR As Long = 2026
Private Const DATE_BOUND_MONTH As Long = 1
Private Const DATE_BOUND_DAY As Long = 1

'VList puts the value at column 5 and the label at column 4.
Private Const VLIST_VALUE_COL As Long = 5
Private Const VLIST_LABEL_COL As Long = 4
'HList puts the variable name on row 8, the main label on row 7 and the first
'data row on row 9.
Private Const HLIST_NAME_ROW As Long = 8
Private Const HLIST_LABEL_ROW As Long = 7
Private Const HLIST_DATA_ROW As Long = 9


'@section Lifecycle
'===============================================================================

'@sub-title Build the fixture workbook once.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestVarWriter"

    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        BuildFixture
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Print results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If

    On Error Resume Next
        If Not FixtureWorkbook Is Nothing Then DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set Dict = Nothing
    Set Specs = Nothing
    Set TargetSheet = Nothing
    Set HListSheet = Nothing
    Set PrintSheet = Nothing
    Set CleanPrintSheet = Nothing
    Set DropStub = Nothing
    Set CustDropStub = Nothing
    Set FixtureWorkbook = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestCreateReturnsAVarWriter()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsAVarWriter"
    If Not FixtureReady("TestCreateReturnsAVarWriter") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsAVarWriter", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestVarWriterRaisesWithoutSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestVarWriterRaisesWithoutSpecs"
    If Not FixtureReady("TestVarWriterRaisesWithoutSpecs") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As VarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Nothing, _
        wksh:=TargetSheet)

    Assert.LogFailure "Create should raise when specs is Nothing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "A missing specs object should raise ObjectNotInitialized - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@TestMethod("VarWriter")
Public Sub TestVarWriterRaisesWithoutWorksheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestVarWriterRaisesWithoutWorksheet"
    If Not FixtureReady("TestVarWriterRaisesWithoutWorksheet") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As VarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=Nothing)

    Assert.LogFailure "Create should raise when wksh is Nothing."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), CLng(Err.Number), _
                    "A missing worksheet should raise ObjectNotInitialized - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@TestMethod("VarWriter")
Public Sub TestCreateRejectsAnUnknownLayer()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsAnUnknownLayer"
    If Not FixtureReady("TestCreateRejectsAnUnknownLayer") Then Exit Sub
    On Error GoTo ExpectError

    Dim sut As VarWriter
    Set sut = VarWriter.Create( _
        layer:=0, _
        specs:=Specs, _
        wksh:=TargetSheet)

    Assert.LogFailure "Create should raise for a layer that names neither member."
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), CLng(Err.Number), _
                    "A layer outside the enum should raise InvalidArgument - description was [" & _
                    Err.Description & "]"
    Err.Clear
End Sub


'@section ValueOf tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestValueOfReadsTheDictionaryRow()
    CustomTestSetTitles Assert, TESTMODULE, "TestValueOfReadsTheDictionaryRow"
    If Not FixtureReady("TestValueOfReadsTheDictionaryRow") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Set sut = NewVListWriter()

    sut.WriteVariable "exp_var_v1"

    Assert.AreEqual "Variable used in export vlist1D", sut.ValueOf("main label"), _
                    "ValueOf should return the dictionary main label value"
    Assert.AreEqual "text", sut.ValueOf("variable type"), _
                    "ValueOf should read every column of the same row"
    Assert.AreEqual vbNullString, sut.ValueOf("no such column"), _
                    "An unknown column should answer an empty string, the way LLVariables does"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueOfReadsTheDictionaryRow", Err.Number, Err.Description
End Sub


'@section VList writing tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestVListWritesLabelToCell()
    CustomTestSetTitles Assert, TESTMODULE, "TestVListWritesLabelToCell"
    If Not FixtureReady("TestVListWritesLabelToCell") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long

    Set sut = NewVListWriter()
    sut.WriteVariable "exp_var_v1"

    rowIdx = ColumnIndexOf("exp_var_v1")

    Assert.IsTrue InStr(1, CStr(TargetSheet.Cells(rowIdx, VLIST_LABEL_COL).Value), _
                  "Variable used in export vlist1D") > 0, _
                  "Label cell should contain the main label text"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVListWritesLabelToCell", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestVListDoesNotWriteToPrint()
    CustomTestSetTitles Assert, TESTMODULE, "TestVListDoesNotWriteToPrint"
    If Not FixtureReady("TestVListDoesNotWriteToPrint") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter

    'CleanPrintSheet is handed to this writer and to no other, so what it holds
    'at the end of the write is what this writer put there.
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        printWksh:=CleanPrintSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "opt_vis_v1"

    Assert.AreEqual CLng(0), CLng(Application.WorksheetFunction.CountA(CleanPrintSheet.UsedRange)), _
                    "VList should not write anything to the print companion sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVListDoesNotWriteToPrint", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestTheFirstVariableWrittenCarriesTheTableAnchor()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheFirstVariableWrittenCarriesTheTableAnchor"
    If Not FixtureReady("TestTheFirstVariableWrittenCarriesTheTableAnchor") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim anchorRng As Range
    Dim tableName As String
    Dim rowIdx As Long

    Set sut = NewVListWriter()

    'mand_v1 sits well past the column index the anchor used to be pinned to, so
    'the anchor exists only because it is the first variable this writer was
    'handed.
    sut.WriteVariable "mand_v1"

    tableName = sut.ValueOf("table name")
    rowIdx = ColumnIndexOf("mand_v1")

    Assert.IsTrue tableName <> vbNullString, _
                  "The prepared dictionary should give the sheet a table name"

    Set anchorRng = TargetSheet.Range(tableName & "_START")

    Assert.AreEqual CLng(rowIdx), CLng(anchorRng.Row), _
                    "The anchor should sit on the row of the first variable written"
    Assert.AreEqual CLng(VLIST_VALUE_COL), CLng(anchorRng.Column), _
                    "The anchor should sit on the VList value column"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheFirstVariableWrittenCarriesTheTableAnchor", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestWritingTheSameVariableTwiceDoesNotRaise()
    CustomTestSetTitles Assert, TESTMODULE, "TestWritingTheSameVariableTwiceDoesNotRaise"
    If Not FixtureReady("TestWritingTheSameVariableTwiceDoesNotRaise") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long

    Set sut = NewVListWriter()

    'date_vali_v1 carries a note, so the second write reaches AddComment on a
    'cell that already has one. Excel raises 1004 for that.
    sut.WriteVariable "date_vali_v1"
    sut.WriteVariable "date_vali_v1"

    rowIdx = ColumnIndexOf("date_vali_v1")

    Assert.IsTrue InStr(1, CStr(TargetSheet.Cells(rowIdx, VLIST_LABEL_COL).Value), _
                  "Date validation on vlist1D") > 0, _
                  "A second write of the same variable should leave the label in place"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestWritingTheSameVariableTwiceDoesNotRaise", Err.Number, Err.Description
End Sub


'@section HList writing tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestHListWritesVarNameToHeader()
    CustomTestSetTitles Assert, TESTMODULE, "TestHListWritesVarNameToHeader"
    If Not FixtureReady("TestHListWritesVarNameToHeader") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim colIdx As Long

    Set sut = NewHListWriter()
    sut.WriteVariable "text_h2"

    colIdx = ColumnIndexOf("text_h2")

    Assert.AreEqual "text_h2", CStr(HListSheet.Cells(HLIST_NAME_ROW, colIdx).Value), _
                    "Variable name should be written to the header row (row 8)"
    Assert.IsTrue InStr(1, CStr(HListSheet.Cells(HLIST_LABEL_ROW, colIdx).Value), _
                  "Random text variable") > 0, _
                  "Main label should be written to the label row (row 7)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHListWritesVarNameToHeader", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestHListWritesToPrintCompanion()
    CustomTestSetTitles Assert, TESTMODULE, "TestHListWritesToPrintCompanion"
    If Not FixtureReady("TestHListWritesToPrintCompanion") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim colIdx As Long

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerHList, _
        specs:=Specs, _
        wksh:=HListSheet, _
        printWksh:=PrintSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "int_h2"

    colIdx = ColumnIndexOf("int_h2")

    Assert.IsTrue InStr(1, CStr(PrintSheet.Cells(HLIST_LABEL_ROW, colIdx).Value), _
                  "Integer on hlist2D") > 0, _
                  "Print sheet should have the main label written"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHListWritesToPrintCompanion", Err.Number, Err.Description
End Sub


'@section Type formatting tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestTextTypeFormatsAsString()
    CustomTestSetTitles Assert, TESTMODULE, "TestTextTypeFormatsAsString"
    If Not FixtureReady("TestTextTypeFormatsAsString") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long

    Set sut = NewVListWriter()

    ' exp_var_v1 has variable type = "text"
    sut.WriteVariable "exp_var_v1"

    rowIdx = ColumnIndexOf("exp_var_v1")

    Assert.AreEqual "@", TargetSheet.Cells(rowIdx, VLIST_VALUE_COL).NumberFormat, _
                    "Text variables should have NumberFormat set to '@'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTextTypeFormatsAsString", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestADecimalFormattedAsTextKeepsItsNumberFormat()
    CustomTestSetTitles Assert, TESTMODULE, "TestADecimalFormattedAsTextKeepsItsNumberFormat"
    If Not FixtureReady("TestADecimalFormattedAsTextKeepsItsNumberFormat") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long
    Dim writtenFormat As String

    Set sut = NewVListWriter()

    'A decimal whose format reads "text" matched no arm of AddType, so the
    'number format was assigned an empty string, which Excel refuses.
    sut.WriteVariable "dec_text_v1"

    rowIdx = ColumnIndexOf("dec_text_v1")
    writtenFormat = CStr(TargetSheet.Cells(rowIdx, VLIST_VALUE_COL).NumberFormat)

    Assert.IsTrue writtenFormat <> vbNullString, _
                  "A decimal formatted as text should keep a number format, found [" & _
                  writtenFormat & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADecimalFormattedAsTextKeepsItsNumberFormat", Err.Number, Err.Description
End Sub


'@section Validation bound tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestADecimalBoundKeepsItsFraction()
    CustomTestSetTitles Assert, TESTMODULE, "TestADecimalBoundKeepsItsFraction"
    If Not FixtureReady("TestADecimalBoundKeepsItsFraction") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long
    Dim boundText As String

    Set sut = NewVListWriter()
    sut.WriteVariable "dec_half_v1"

    rowIdx = ColumnIndexOf("dec_half_v1")
    boundText = ValidationBound(TargetSheet.Cells(rowIdx, VLIST_VALUE_COL))

    'CLng turned a maximum of 0.5 into 0, and the validation then rejected every
    'value the dictionary meant to allow.
    Assert.IsTrue boundText <> vbNullString, _
                  "A decimal maximum should produce a validation - max reads [" & _
                  sut.ValueOf("max") & "], type reads [" & sut.ValueOf("variable type") & _
                  "], the writer filed [" & FirstCheckingLabel(sut) & "]"
    Assert.IsTrue BoundAsNumber(boundText) > 0, _
                  "A decimal maximum of 0.5 should not be truncated to zero, found [" & _
                  boundText & "]"
    Assert.IsTrue Abs(BoundAsNumber(boundText) - 0.5) < 0.000001, _
                  "The decimal maximum should read 0.5, found [" & boundText & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADecimalBoundKeepsItsFraction", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestAWholeNumberBoundAbove32767KeepsItsValidation()
    CustomTestSetTitles Assert, TESTMODULE, "TestAWholeNumberBoundAbove32767KeepsItsValidation"
    If Not FixtureReady("TestAWholeNumberBoundAbove32767KeepsItsValidation") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long
    Dim boundText As String

    Set sut = NewVListWriter()
    sut.WriteVariable "big_int_v1"

    rowIdx = ColumnIndexOf("big_int_v1")

    'CInt raised error 6 past 32,767 and AddValidation runs under a suppression,
    'so the variable used to end up with no validation and no message.
    boundText = ValidationBound(TargetSheet.Cells(rowIdx, VLIST_VALUE_COL))

    Assert.IsTrue boundText <> vbNullString, _
                  "A whole-number maximum should produce a validation - the writer filed [" & _
                  FirstCheckingLabel(sut) & "]"
    Assert.IsTrue Abs(BoundAsNumber(boundText) - 100000) < 0.000001, _
                  "A whole-number maximum of 100000 should survive the conversion, found [" & _
                  boundText & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAWholeNumberBoundAbove32767KeepsItsValidation", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestADateMinFormattedAsADateKeepsItsValidation()
    CustomTestSetTitles Assert, TESTMODULE, "TestADateMinFormattedAsADateKeepsItsValidation"
    If Not FixtureReady("TestADateMinFormattedAsADateKeepsItsValidation") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long
    Dim boundText As String
    Dim cellRng As Range
    Dim readBound As String

    Set sut = NewVListWriter()
    sut.WriteVariable DATE_BOUND_VAR

    rowIdx = ColumnIndexOf(DATE_BOUND_VAR)
    Set cellRng = TargetSheet.Cells(rowIdx, VLIST_VALUE_COL)
    boundText = ValidationBound(cellRng)

    'What VarWriter read out of the dictionary, quoted into every message: a
    'date cell reaches ValueOf through CStr, so this is the string the parser
    'and CDate are handed.
    readBound = sut.ValueOf("min")

    Assert.IsTrue boundText <> vbNullString, _
                  "A date minimum should produce a validation - min reads [" & readBound & _
                  "], type reads [" & sut.ValueOf("variable type") & _
                  "], the writer filed [" & FirstCheckingLabel(sut) & "]"

    Assert.AreEqual CLng(xlValidateDate), CLng(ValidationType(cellRng)), _
                    "A date minimum should make a date validation, min reads [" & readBound & _
                    "], bound reads [" & boundText & "]"

    Assert.IsTrue (BoundAsDate(boundText) = DATE_BOUND_MIN()), _
                  "The date minimum should reach the validation as " & _
                  Format$(DATE_BOUND_MIN(), "yyyy-mm-dd") & ", min reads [" & readBound & _
                  "], bound reads [" & boundText & "], which is " & _
                  Format$(BoundAsDate(boundText), "yyyy-mm-dd")

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestADateMinFormattedAsADateKeepsItsValidation", Err.Number, Err.Description
End Sub


'@section Formula tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestARejectedFormulaLeavesTheCellAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestARejectedFormulaLeavesTheCellAlone"
    If Not FixtureReady("TestARejectedFormulaLeavesTheCellAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long
    Dim cellRng As Range

    Set sut = NewVListWriter()

    'brok_form_v1 names a variable the dictionary does not carry, so the parser
    'rejects it.
    sut.WriteVariable "brok_form_v1"

    rowIdx = ColumnIndexOf("brok_form_v1")
    Set cellRng = TargetSheet.Cells(rowIdx, VLIST_VALUE_COL)

    Assert.IsTrue Not cellRng.HasFormula, _
                  "A rejected formula should leave no formula on the cell"
    Assert.IsTrue Not cellRng.Locked, _
                  "A cell with no formula should stay unlocked, so what the user sees " & _
                  "and what the sheet allows agree"
    Assert.IsTrue sut.HasCheckings, _
                  "A rejected formula should be filed through Checking"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARejectedFormulaLeavesTheCellAlone", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestAParsedFormulaLocksItsCell()
    CustomTestSetTitles Assert, TESTMODULE, "TestAParsedFormulaLocksItsCell"
    If Not FixtureReady("TestAParsedFormulaLocksItsCell") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long
    Dim cellRng As Range

    Set sut = NewVListWriter()

    'form_v1 names choi_v1, which the dictionary does carry.
    sut.WriteVariable "form_v1"

    rowIdx = ColumnIndexOf("form_v1")
    Set cellRng = TargetSheet.Cells(rowIdx, VLIST_VALUE_COL)

    Assert.IsTrue cellRng.HasFormula, _
                  "A parsed formula should reach the cell"
    Assert.IsTrue cellRng.Locked, _
                  "A calculated cell should be locked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAParsedFormulaLocksItsCell", Err.Number, Err.Description
End Sub


'@section Choice tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestAChoiceVariableGetsItsDropdown()
    CustomTestSetTitles Assert, TESTMODULE, "TestAChoiceVariableGetsItsDropdown"
    If Not FixtureReady("TestAChoiceVariableGetsItsDropdown") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim rowIdx As Long

    Set sut = NewVListWriter()
    sut.WriteVariable "choi_v1"

    rowIdx = ColumnIndexOf("choi_v1")

    Assert.AreEqual CLng(xlValidateList), _
                    CLng(TargetSheet.Cells(rowIdx, VLIST_VALUE_COL).Validation.Type), _
                    "A choice_manual variable should carry a list validation"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAChoiceVariableGetsItsDropdown", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestAWriterWithoutDropdownsFilesAChecking()
    CustomTestSetTitles Assert, TESTMODULE, "TestAWriterWithoutDropdownsFilesAChecking"
    If Not FixtureReady("TestAWriterWithoutDropdownsFilesAChecking") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet)

    'The two managers are optional at creation, and AddChoices used to call
    'SetValidation on whichever of them the control type asked for.
    sut.WriteVariable "choi_ord_v1"

    Assert.IsTrue sut.HasCheckings, _
                  "A choice variable with no dropdown manager should be filed, not raised"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAWriterWithoutDropdownsFilesAChecking", Err.Number, Err.Description
End Sub


'@section Milestone tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestAWrittenVariableFilesOneMilestone()
    CustomTestSetTitles Assert, TESTMODULE, "TestAWrittenVariableFilesOneMilestone"
    If Not FixtureReady("TestAWrittenVariableFilesOneMilestone") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter

    Set sut = NewVListWriter()

    Assert.IsFalse sut.HasMilestones, _
                   "A writer that has written nothing should carry no milestone"

    sut.WriteVariable "exp_var_v1"

    Assert.IsTrue sut.HasMilestones, "A written variable should file a milestone"
    Assert.AreEqual CLng(1), CLng(sut.MilestoneValues.Length), _
                    "One written variable is one milestone entry"
    Assert.AreEqual CLng(1), sut.VariablesWritten, _
                    "One written variable is a count of one"
    Assert.AreEqual "exp_var_v1 on " & TargetSheet.Name, _
                    sut.MilestoneValues.ValueOf("exp_var_v1", checkingLabel), _
                    "The milestone is keyed by the variable and labelled with its sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAWrittenVariableFilesOneMilestone", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestMilestonesStayOutOfTheCheckings()
    CustomTestSetTitles Assert, TESTMODULE, "TestMilestonesStayOutOfTheCheckings"
    If Not FixtureReady("TestMilestonesStayOutOfTheCheckings") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter

    'The two stores travel to different places: the problems reach the
    '__check worksheet, the milestones reach the text record alone. A
    'milestone leaking into the problems is what puts an EntireColumn.AutoFit
    'behind every variable of the build.
    Set sut = NewVListWriter()
    sut.WriteVariable "exp_var_v1"

    Assert.IsFalse sut.HasCheckings, _
                   "A variable that wrote cleanly should file no problem entry"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestMilestonesStayOutOfTheCheckings", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestAVariableWithoutAColumnIndexIsLeftOutOfTheCount()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVariableWithoutAColumnIndexIsLeftOutOfTheCount"
    If Not FixtureReady("TestAVariableWithoutAColumnIndexIsLeftOutOfTheCount") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter

    Set sut = NewVListWriter()

    'A name the dictionary does not carry resolves to no cell, so the writer
    'files the problem and leaves the sheet alone.
    sut.WriteVariable "no_such_variable_at_all"

    Assert.AreEqual CLng(0), sut.VariablesWritten, _
                    "A variable that was never written should not be counted"
    Assert.IsFalse sut.HasMilestones, _
                   "A variable that was never written should file no milestone"
    Assert.IsTrue sut.HasCheckings, _
                  "A variable that could not be placed should be filed as a problem"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVariableWithoutAColumnIndexIsLeftOutOfTheCount", Err.Number, Err.Description
End Sub


'@section Conditional formatting tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestAConditionalFormatVariableGetsOneCondition()
    CustomTestSetTitles Assert, TESTMODULE, "TestAConditionalFormatVariableGetsOneCondition"
    If Not FixtureReady("TestAConditionalFormatVariableGetsOneCondition") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As VarWriter
    Dim colIdx As Long
    Dim cellRng As Range

    Set sut = NewHListWriter()

    'cond_val_h1 points its formatting condition at cond_test_h1, and both live
    'on this sheet.
    sut.WriteVariable "cond_val_h1"

    colIdx = ColumnIndexOf("cond_val_h1")
    Set cellRng = HListSheet.Cells(HLIST_DATA_ROW, colIdx)

    'Two calls to FormatConditions.Add both took, so the cell used to end with
    'two conditions and only the last of them was given a colour.
    Assert.AreEqual CLng(1), CLng(cellRng.FormatConditions.Count), _
                    "A conditional format variable should leave exactly one condition"

    'A count on its own cannot tell the fault this test was written for from
    'Excel declining both formulas. AddOneCondition files a warning naming the
    'variable when it is refused, so reading it labels the next occurrence
    'instead of leaving a bare count to be guessed at. This answered
    'expected 1, actual 0 on 2026-08-01 and again on 2026-08-11, both times on
    'a tree whose diff reached nothing in this folder.
    Assert.IsFalse sut.HasCheckings, _
                   "The writer should file nothing while adding one condition - " & _
                   "a checking here means Excel refused both formulas"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAConditionalFormatVariableGetsOneCondition", Err.Number, Err.Description
End Sub


'@section Fixture
'===============================================================================

'@sub-title Build the workbook, the dictionary and the specifications.
Private Sub BuildFixture()
    Dim transStub As TranslationObject
    Dim design As LLFormat
    Dim formatSheet As Worksheet
    Dim formData As FormulaData
    Dim formulaSheet As Worksheet
    Dim choicesObj As LLChoices

    Set FixtureWorkbook = NewWorkbook
    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, FixtureWorkbook
    SeedBoundaryVariables FixtureWorkbook.Worksheets(DICTIONARY_SHEET)
    Set Dict = LLdictionary.Create(FixtureWorkbook.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare

    ChoicesTestFixture.PrepareChoicesFixture CHOICES_SHEET, FixtureWorkbook
    Set choicesObj = LLChoices.Create(FixtureWorkbook.Worksheets(CHOICES_SHEET), 1, 1)

    'The three MSG_ tags AddLabel reads live in the table, the way a real
    'linelist translation table carries them.
    Set transStub = BuildTranslationObject(FixtureWorkbook, "ENG", _
        Array(Array("MSG_Calculated", "Calculated"), _
              Array("MSG_Mandatory", "Mandatory"), _
              Array("MSG_CustomChoice", "Custom choice")))
    Set formatSheet = LLFormatTestFixture.PrepareLLFormatFixture("LLFormatFixture", FixtureWorkbook)
    Set design = LLFormat.Create(formatSheet)
    Set formulaSheet = FormulaTestFixture.PrepareFormulaFixtureSheet("FormulaFixture", outwb:=FixtureWorkbook)
    Set formData = FormulaData.Create(formulaSheet)

    Set TargetSheet = FixtureWorkbook.Worksheets.Add
    TargetSheet.Name = VLIST_SHEET
    Set HListSheet = FixtureWorkbook.Worksheets.Add
    HListSheet.Name = HLIST_SHEET
    Set PrintSheet = FixtureWorkbook.Worksheets.Add
    Set CleanPrintSheet = FixtureWorkbook.Worksheets.Add

    'Dropdown lists live on their own host sheets (as in production) so their
    'list tables and validations never collide with the variable content the
    'tests assert on TargetSheet.
    Set DropStub = DropdownLists.Create(FixtureWorkbook.Worksheets.Add, hprefix:=vbNullString)
    Set CustDropStub = DropdownLists.Create(FixtureWorkbook.Worksheets.Add, hprefix:=vbNullString)

    EnsureSpecsSheets FixtureWorkbook
    Set Specs = LinelistSpecs.Create(FixtureWorkbook)
    Specs.TestAssignDictionary Dict
    Specs.TestAssignDesignFormat design
    Specs.TestAssignTransObject transStub
    Specs.TestAssignFormulaData formData
    Specs.TestAssignChoices choicesObj
End Sub

'@fun-title Report a fixture that could not be built, once per test.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is there.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If SetupError = 0 Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function

'@description Append the three variables the boundary tests need.
'@details The shared dictionary fixture is fixed, so the rows a test needs and
'the fixture does not carry are written onto the fixture sheet before
'LLdictionary reads it. They join vlist1D-sheet1 with no section, which is what
'the last rows of that sheet already look like.
Private Sub SeedBoundaryVariables(ByVal dictSheet As Worksheet)
    Dim firstFreeRow As Long

    firstFreeRow = DictionaryTestFixture.DictionaryFixtureRowCount() + 2

    'The bound carries the host's decimal separator. Formulas tokenises a
    'decimal literal with that separator alone, so a bound written the other way
    'never reaches the conversion this test is about.
    SeedVariable dictSheet, firstFreeRow, "dec_half_v1", "Decimal bounded at a half", _
                 "decimal", vbNullString, "0" & HostDecimalSeparator() & "5"
    SeedVariable dictSheet, firstFreeRow + 1, "big_int_v1", "Integer above the Integer type", _
                 "integer", vbNullString, "100000"
    SeedVariable dictSheet, firstFreeRow + 2, "dec_text_v1", "Decimal formatted as text", _
                 "decimal", "text", vbNullString
    SeedDateBoundVariable dictSheet, firstFreeRow + 3
End Sub

'@description Write a variable whose Min cell is a REAL DATE carrying a date format.
'@details
'Every other seeded row writes its bound into a cell forced to text, so the
'string the dictionary carries is the string the author typed. A dictionary
'authored in Excel does not have to look like that: a min or max meant as a date
'is typed into a cell, Excel stores a Date and formats it, and what VarWriter
'reads back is whatever CStr makes of that Date on this host. This row is that
'case, and it is the only one in the suite.
Private Sub SeedDateBoundVariable(ByVal dictSheet As Worksheet, ByVal rowNumber As Long)

    SeedCell dictSheet, rowNumber, "Variable Name", DATE_BOUND_VAR
    SeedCell dictSheet, rowNumber, "Main Label", "Date bounded by a formatted date cell"
    SeedCell dictSheet, rowNumber, "Sheet Name", VLIST_SHEET
    SeedCell dictSheet, rowNumber, "Sheet Type", "vlist1D"
    SeedCell dictSheet, rowNumber, "Variable Type", "date"
    SeedCell dictSheet, rowNumber, "Alert", "error"
    SeedCell dictSheet, rowNumber, "Message", "Seeded for the date bound test"

    'The point of the row: a Date value under a date number format, which is what
    'a person typing a date into the dictionary leaves behind.
    SeedDateCell dictSheet, rowNumber, "Min", DATE_BOUND_MIN
End Sub

'@description Write one fixture cell as a real Date under a date number format.
Private Sub SeedDateCell(ByVal dictSheet As Worksheet, _
                         ByVal rowNumber As Long, _
                         ByVal headerName As String, _
                         ByVal cellValue As Date)
    Dim columnNumber As Long

    columnNumber = DictionaryTestFixture.DictionaryHeaderIndex(headerName) + 1

    dictSheet.Cells(rowNumber, columnNumber).NumberFormat = "dd/mm/yyyy"
    dictSheet.Cells(rowNumber, columnNumber).Value = cellValue
End Sub

'@description Write one dictionary row by header name.
Private Sub SeedVariable(ByVal dictSheet As Worksheet, _
                         ByVal rowNumber As Long, _
                         ByVal varName As String, _
                         ByVal mainLabel As String, _
                         ByVal varType As String, _
                         ByVal varFormat As String, _
                         ByVal maxValue As String)

    SeedCell dictSheet, rowNumber, "Variable Name", varName
    SeedCell dictSheet, rowNumber, "Main Label", mainLabel
    SeedCell dictSheet, rowNumber, "Sheet Name", VLIST_SHEET
    SeedCell dictSheet, rowNumber, "Sheet Type", "vlist1D"
    SeedCell dictSheet, rowNumber, "Variable Type", varType
    SeedCell dictSheet, rowNumber, "Variable Format", varFormat
    SeedCell dictSheet, rowNumber, "Max", maxValue
    SeedCell dictSheet, rowNumber, "Alert", "error"
    SeedCell dictSheet, rowNumber, "Message", "Seeded for the boundary tests"
End Sub

'@description Write one cell of the fixture sheet, located by its header name.
Private Sub SeedCell(ByVal dictSheet As Worksheet, _
                     ByVal rowNumber As Long, _
                     ByVal headerName As String, _
                     ByVal cellValue As String)
    Dim columnNumber As Long

    'DictionaryHeaderIndex answers a zero-based offset into the header array,
    'and the headers are written from column 1.
    columnNumber = DictionaryTestFixture.DictionaryHeaderIndex(headerName) + 1

    'The cell is made text first. A bound written as "0.5" into a General cell
    'is stored as the NUMBER 0.5, and reading it back through CStr renders it
    'with the host's decimal separator -- which is a different string from the
    'one a dictionary authored elsewhere carries.
    dictSheet.Cells(rowNumber, columnNumber).NumberFormat = "@"
    dictSheet.Cells(rowNumber, columnNumber).Value = cellValue
End Sub


'@section Helpers
'===============================================================================

'@description Build a VList writer over the VList sheet with both dropdown managers.
Private Function NewVListWriter() As VarWriter
    Set NewVListWriter = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)
End Function

'@description Build an HList writer over the HList sheet with both dropdown managers.
Private Function NewHListWriter() As VarWriter
    Set NewHListWriter = VarWriter.Create( _
        layer:=VarWriterLayerHList, _
        specs:=Specs, _
        wksh:=HListSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)
End Function

'@description The day the date-bound row carries in its Min cell.
'@details A Const cannot hold a DateSerial call, so the one day the seed and the
'assertion both name lives here.
Private Function DATE_BOUND_MIN() As Date
    DATE_BOUND_MIN = DateSerial(DATE_BOUND_YEAR, DATE_BOUND_MONTH, DATE_BOUND_DAY)
End Function

'@description Read a validation bound as a date, whatever shape Excel stored it in.
'@details
'Validation.Formula1 comes back as text. A date bound can arrive as a date
'string the host can read, or as the serial number behind it, so both are
'tried. Answers 0 when it is neither, which is what a bound that came through
'as arithmetic looks like.
Private Function BoundAsDate(ByVal boundText As String) As Date
    Dim stripped As String
    Dim serialValue As Double

    If LenB(boundText) = 0 Then Exit Function

    stripped = Trim$(Replace(boundText, "=", vbNullString))
    If LenB(stripped) = 0 Then Exit Function

    On Error Resume Next
        BoundAsDate = CDate(stripped)
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
    On Error GoTo 0

    serialValue = BoundAsNumber(stripped)
    If serialValue < 1 Then Exit Function

    On Error Resume Next
        BoundAsDate = CDate(serialValue)
        Err.Clear
    On Error GoTo 0
End Function

'@description The decimal separator this host writes and reads numbers with.
Private Function HostDecimalSeparator() As String
    HostDecimalSeparator = Mid$(CStr(1.5), 2, 1)
End Function

'@description Read a validation bound as a number, whichever separator it carries.
Private Function BoundAsNumber(ByVal boundText As String) As Double
    Dim normalised As String

    If LenB(boundText) = 0 Then Exit Function

    normalised = Replace(boundText, ".", HostDecimalSeparator())
    normalised = Replace(normalised, ",", HostDecimalSeparator())

    On Error Resume Next
        BoundAsNumber = CDbl(normalised)
        Err.Clear
    On Error GoTo 0
End Function

'@description Read a cell's validation type, answering 0 when it carries none.
Private Function ValidationType(ByVal cellRng As Range) As Long
    On Error Resume Next
        ValidationType = cellRng.Validation.Type
        Err.Clear
    On Error GoTo 0
End Function

'@description Read a cell's validation bound, answering empty when it carries none.
Private Function ValidationBound(ByVal cellRng As Range) As String
    On Error Resume Next
        ValidationBound = CStr(cellRng.Validation.Formula1)
        Err.Clear
    On Error GoTo 0
End Function

'@description Read the first entry a writer filed, so a failure message carries it.
Private Function FirstCheckingLabel(ByVal writer As VarWriter) As String
    Dim keys As BetterArray

    FirstCheckingLabel = "no checking was filed"
    If writer Is Nothing Then Exit Function
    If Not writer.HasCheckings Then Exit Function

    Set keys = writer.CheckingValues.ListOfKeys()
    If keys Is Nothing Then Exit Function
    If keys.Length = 0 Then Exit Function

    FirstCheckingLabel = writer.CheckingValues.ValueOf(CStr(keys.Item(keys.LowerBound)), checkingLabel)
End Function

'@description Read the column index Dict.Prepare assigned to a variable.
Private Function ColumnIndexOf(ByVal varName As String) As Long
    Dim vars As LLVariables

    Set vars = LLVariables.Create(Dict)
    ColumnIndexOf = CLng(vars.Value("column index", varName))
End Function
