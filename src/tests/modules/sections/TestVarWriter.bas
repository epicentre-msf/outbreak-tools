Attribute VB_Name = "TestVarWriter"
Attribute VB_Description = "Tests for VarWriter class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for VarWriter class")

Option Explicit

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private Dict As ILLdictionary
Private Specs As LLVarContextSpecsStub
Private TargetSheet As Worksheet
Private PrintSheet As Worksheet
Private DropStub As DropdownListsStub
Private CustDropStub As DropdownListsStub

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "VarWriter"
Private Const DICTIONARY_SHEET As String = "DictFixture"


'@section Lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    TestHelpers.EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestVarWriter"
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
    Set PrintSheet = FixtureWorkbook.Worksheets.Add

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
    Set PrintSheet = Nothing
    Set DropStub = Nothing
    Set CustDropStub = Nothing
    Set FixtureWorkbook = Nothing
End Sub


'@section Factory tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestCreateReturnsInstance()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateReturnsInstance"
    On Error GoTo TestFail

    Dim sut As IVarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create should return a non-Nothing instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateReturnsInstance", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestCreateRaisesWithoutSpecs()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRaisesWithoutSpecs"
    On Error GoTo ExpectError

    Dim sut As IVarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Nothing, _
        wksh:=TargetSheet)

    Assert.LogFailure "Create should raise when specs is Nothing."
    Exit Sub

ExpectError:
    Assert.IsTrue Err.Number <> 0, "An error should have been raised"
    Err.Clear
End Sub


'@TestMethod("VarWriter")
Public Sub TestCreateRaisesWithoutWorksheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRaisesWithoutWorksheet"
    On Error GoTo ExpectError

    Dim sut As IVarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=Nothing)

    Assert.LogFailure "Create should raise when wksh is Nothing."
    Exit Sub

ExpectError:
    Assert.IsTrue Err.Number <> 0, "An error should have been raised"
    Err.Clear
End Sub


'@section ValueOf tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestValueOfReadsMainLabel()
    CustomTestSetTitles Assert, TESTMODULE, "TestValueOfReadsMainLabel"
    On Error GoTo TestFail

    Dim sut As IVarWriter
    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "exp_var_v1"

    Assert.AreEqual "Variable used in export vlist1D", sut.ValueOf("main label"), _
                     "ValueOf should return the dictionary main label value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueOfReadsMainLabel", Err.Number, Err.Description
End Sub


'@section VList writing tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestVListWritesLabelToCell()
    CustomTestSetTitles Assert, TESTMODULE, "TestVListWritesLabelToCell"
    On Error GoTo TestFail

    Dim sut As IVarWriter
    Dim vars As ILLVariables
    Dim colIdx As Long

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "exp_var_v1"

    ' Look up the column index assigned by Dict.Prepare
    Set vars = LLVariables.Create(Dict)
    colIdx = CLng(vars.Value("column index", "exp_var_v1"))

    ' VList label is written to Offset(,-1) of VarRange = Cells(colIdx, 5)
    ' So label is at Cells(colIdx, 4)
    Assert.IsTrue InStr(1, CStr(TargetSheet.Cells(colIdx, 4).Value), _
                  "Variable used in export vlist1D") > 0, _
                  "Label cell should contain the main label text"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVListWritesLabelToCell", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestVListDoesNotWriteToPrint()
    CustomTestSetTitles Assert, TESTMODULE, "TestVListDoesNotWriteToPrint"
    On Error GoTo TestFail

    Dim sut As IVarWriter

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        printWksh:=PrintSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "exp_var_v1"

    ' VList should NOT write to the print companion
    Assert.AreEqual CLng(0), CLng(Application.WorksheetFunction.CountA(PrintSheet.UsedRange)), _
                     "VList should not write anything to the print companion sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestVListDoesNotWriteToPrint", Err.Number, Err.Description
End Sub


'@section HList writing tests
'===============================================================================

'@TestMethod("VarWriter")
Public Sub TestHListWritesVarNameToHeader()
    CustomTestSetTitles Assert, TESTMODULE, "TestHListWritesVarNameToHeader"
    On Error GoTo TestFail

    Dim sut As IVarWriter
    Dim vars As ILLVariables
    Dim colIdx As Long

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerHList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "text_h2"

    ' Look up the column index assigned by Dict.Prepare
    Set vars = LLVariables.Create(Dict)
    colIdx = CLng(vars.Value("column index", "text_h2"))

    ' HList writes var name at VarRange.Offset(-1) = Cells(8, colIdx)
    Assert.AreEqual "text_h2", CStr(TargetSheet.Cells(8, colIdx).Value), _
                     "Variable name should be written to the header row (row 8)"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestHListWritesVarNameToHeader", Err.Number, Err.Description
End Sub


'@TestMethod("VarWriter")
Public Sub TestHListWritesToPrintCompanion()
    CustomTestSetTitles Assert, TESTMODULE, "TestHListWritesToPrintCompanion"
    On Error GoTo TestFail

    Dim sut As IVarWriter
    Dim vars As ILLVariables
    Dim colIdx As Long

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerHList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        printWksh:=PrintSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    sut.WriteVariable "text_h2"

    ' Look up the column index assigned by Dict.Prepare
    Set vars = LLVariables.Create(Dict)
    colIdx = CLng(vars.Value("column index", "text_h2"))

    ' HList writes the main label to the print companion at Offset(-2) = row 7
    Assert.IsTrue InStr(1, CStr(PrintSheet.Cells(7, colIdx).Value), _
                  "Random text variable") > 0, _
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
    On Error GoTo TestFail

    Dim sut As IVarWriter
    Dim vars As ILLVariables
    Dim colIdx As Long

    Set sut = VarWriter.Create( _
        layer:=VarWriterLayerVList, _
        specs:=Specs, _
        wksh:=TargetSheet, _
        dropdownObj:=DropStub, _
        customDropdownObj:=CustDropStub)

    ' exp_var_v1 has variable type = "text"
    sut.WriteVariable "exp_var_v1"

    Set vars = LLVariables.Create(Dict)
    colIdx = CLng(vars.Value("column index", "exp_var_v1"))

    ' VList VarRange = Cells(colIdx, 5). AddType sets NumberFormat = "@" for text
    Assert.AreEqual "@", TargetSheet.Cells(colIdx, 5).NumberFormat, _
                     "Text variables should have NumberFormat set to '@'"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTextTypeFormatsAsString", Err.Number, Err.Description
End Sub
