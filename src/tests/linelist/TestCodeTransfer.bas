Attribute VB_Name = "TestCodeTransfer"
Attribute VB_Description = "Tests for CodeTransfer class"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for CodeTransfer class")

Option Explicit

'@description
'Drives CodeTransfer, which moves the VBA components of a generated linelist
'from the designer workbook into the output workbook. Linelist.TransferAllCode
'is its one consumer and it moves 39 components per generation, so what this
'suite watches is whether the delivered file compiles.
'
'THE FIXTURE ADDS ITS OWN COMPONENTS
'-------------------------------------------------------------------------------
'The nine forms this class transfers live inside the designer workbook and no
'copy of them is in src/, so a form cannot be given a fixture here. What the
'fixture does instead is add a standard module and a class module to the source
'workbook through VBComponents.Add and write one known comment line into each.
'The tests then read that line back out of the target.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'Adding a component needs "Trust access to the VBA project object model". The
'headless runner needs the same setting to import its own modules, so it is on
'during a run. When it is off, an error escaping TestInitialize would open a
'modal dialog in the VBE and stop the whole run. The setup captures the error
'instead and each test reports it as its own failure.
'
'A TRANSFER REPLACES WHAT THE TARGET HOLDS
'-------------------------------------------------------------------------------
'Owner decision of 2026-07-31: a component already in the target is removed
'before the import and the removal is filed through Checking. Two tests state
'both halves of it.
'@depends CodeTransfer, TemporaryRepos, CustomTest

Private Assert As CustomTest
Private SourceWkb As Workbook
Private TargetWkb As Workbook
Private TempRepos As TemporaryRepos
Private BaseFolder As String
Private SetupError As Long
Private SetupMessage As String

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "CodeTransfer"

Private Const PROBE_MODULE As String = "OBTProbeModule"
Private Const PROBE_CLASS As String = "OBTProbeClass"
Private Const PROBE_ABSENT As String = "OBTProbeAbsent"

Private Const SOURCE_LINE As String = "'obt probe written in the source workbook"
Private Const SOURCE_MARK As String = "written in the source workbook"
Private Const TARGET_LINE As String = "'obt probe already sitting in the target"
Private Const TARGET_MARK As String = "already sitting in the target"

'VBComponents.Add takes a vbext_ComponentType. The project holds no reference to
'the VBA Extensibility library, so the two values are written out here.
Private Const VB_STANDARD_MODULE As Long = 1
Private Const VB_CLASS_MODULE As Long = 2


'@section Lifecycle
'===============================================================================

'@sub-title Set up the assertion harness and a writable base folder.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestCodeTransfer"

    'The repository roots itself under this folder. The default base is
    'Application.DefaultFilePath, which sits outside the folder the run is
    'granted, and a write there is the sandbox prompt that shows up as -1712.
    BaseFolder = BuildTempFolder(Nothing, vbNullString)

    'Two tests count the files in the repository. A file an earlier run left
    'there would be counted with them, so the folder is emptied once.
    Dim sweeper As TemporaryRepos
    On Error Resume Next
        Set sweeper = TemporaryRepos.Create(BaseFolder)
        sweeper.DeleteAll
    On Error GoTo 0
End Sub

'@sub-title Print results, empty the repository and tear down.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'
'The repository folder sits directly under the run directory, and the suite
'used to leave it there. An empty OBTApp_ folder in the runner output is what
'this removes.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RemoveRepository

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Delete the repository files and the folder holding them.
Private Sub RemoveRepository()
    Dim sweeper As TemporaryRepos
    Dim root As String

    If LenB(BaseFolder) = 0 Then Exit Sub

    On Error Resume Next
        Set sweeper = TemporaryRepos.Create(BaseFolder)
        root = sweeper.RootPath
        sweeper.DeleteAll

        'RootPath ends with a path separator and RmDir wants the folder itself.
        If Right$(root, 1) = Application.PathSeparator Then
            root = Left$(root, Len(root) - 1)
        End If
        If LenB(root) > 0 Then RmDir root
    On Error GoTo 0
End Sub

'@sub-title Build two workbooks, a repository and the two probe components.
'@details
'There is no BeginTest call here on purpose. BeginTest opens the checking with
'whatever titles are pending at that moment, and the Flush in TestCleanup has
'just reset those to the default, so every result of the module would be filed
'under the default label. Letting the first assertion of each test open the
'checking picks up the titles CustomTestSetTitles set at the top of the test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    SetupError = 0
    SetupMessage = vbNullString

    On Error Resume Next
        Set SourceWkb = NewWorkbook()
        Set TargetWkb = NewWorkbook()
        Set TempRepos = TemporaryRepos.Create(BaseFolder)
        AddComponent SourceWkb, PROBE_MODULE, VB_STANDARD_MODULE, SOURCE_LINE
        AddComponent SourceWkb, PROBE_CLASS, VB_CLASS_MODULE, SOURCE_LINE
        SetupError = Err.Number
        SetupMessage = Err.Description
    On Error GoTo 0
End Sub

'@sub-title Flush the results and drop everything the test wrote.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        If Not TempRepos Is Nothing Then TempRepos.Reset
        If Not SourceWkb Is Nothing Then DeleteWorkbook SourceWkb
        If Not TargetWkb Is Nothing Then DeleteWorkbook TargetWkb
    On Error GoTo 0

    Set SourceWkb = Nothing
    Set TargetWkb = Nothing
    Set TempRepos = Nothing
End Sub


'@section Fixture helpers
'===============================================================================

'@sub-title Add a component carrying one known line to a workbook.
'@param book Workbook. The workbook to add the component to.
'@param componentName String. The name to give the new component.
'@param componentKind Long. 1 for a standard module, 2 for a class module.
'@param codeLine String. The line written into the new component.
Private Sub AddComponent(ByVal book As Workbook, _
                         ByVal componentName As String, _
                         ByVal componentKind As Long, _
                         ByVal codeLine As String)
    Dim vbComp As Object

    Set vbComp = book.VBProject.VBComponents.Add(componentKind)
    vbComp.Name = componentName
    vbComp.CodeModule.AddFromString codeLine
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

'@fun-title Count the components of a workbook carrying a given name.
'@param book Workbook. The workbook to walk.
'@param componentName String. The name to count.
'@return Long. How many components answer to that name.
Private Function ComponentCount(ByVal book As Workbook, _
                                ByVal componentName As String) As Long
    Dim vbComp As Object
    Dim total As Long

    For Each vbComp In book.VBProject.VBComponents
        If StrComp(vbComp.Name, componentName, vbTextCompare) = 0 Then
            total = total + 1
        End If
    Next

    ComponentCount = total
End Function

'@fun-title Read the whole code of one component.
'@param book Workbook. The workbook holding the component.
'@param componentName String. The component to read.
'@return String. Every line of it, and an empty string when it holds none.
Private Function ComponentCode(ByVal book As Workbook, _
                               ByVal componentName As String) As String
    With book.VBProject.VBComponents(componentName).CodeModule
        If .CountOfLines > 0 Then ComponentCode = .Lines(1, .CountOfLines)
    End With
End Function

'@fun-title Count the files sitting in a folder.
'@param folderPath String. The folder to walk, separator included.
'@return Long. How many files are in it.
Private Function FileCountIn(ByVal folderPath As String) As Long
    Dim fileName As String
    Dim total As Long

    On Error Resume Next
        fileName = Dir$(folderPath & "*")
        Do While LenB(fileName) > 0
            total = total + 1
            fileName = Dir$
        Loop
    On Error GoTo 0

    FileCountIn = total
End Function


'@section Factory tests
'===============================================================================

'@sub-title Three live arguments give an instance.
'@TestMethod("CodeTransfer")
Public Sub TestCreateWiresTheThreeArguments()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateWiresTheThreeArguments"
    If Not FixtureReady("TestCreateWiresTheThreeArguments") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    Assert.IsTrue Not sut Is Nothing, _
                  "Create gives back an instance"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateWiresTheThreeArguments", Err.Number, Err.Description
End Sub

'@sub-title A Nothing source workbook is refused, and the number says why.
'@TestMethod("CodeTransfer")
Public Sub TestCreateRejectsNothingSource()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingSource"
    If Not FixtureReady("TestCreateRejectsNothingSource") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    On Error Resume Next
        Set sut = CodeTransfer.Create(Nothing, TargetWkb, TempRepos)
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing source workbook is refused - description was [" & _
                    errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingSource", Err.Number, Err.Description
End Sub

'@sub-title A Nothing target workbook is refused, and the number says why.
'@TestMethod("CodeTransfer")
Public Sub TestCreateRejectsNothingTarget()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingTarget"
    If Not FixtureReady("TestCreateRejectsNothingTarget") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    On Error Resume Next
        Set sut = CodeTransfer.Create(SourceWkb, Nothing, TempRepos)
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing target workbook is refused - description was [" & _
                    errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTarget", Err.Number, Err.Description
End Sub

'@sub-title A Nothing repository is refused, and the number says why.
'@TestMethod("CodeTransfer")
Public Sub TestCreateRejectsNothingTempRepos()
    CustomTestSetTitles Assert, TESTMODULE, "TestCreateRejectsNothingTempRepos"
    If Not FixtureReady("TestCreateRejectsNothingTempRepos") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    On Error Resume Next
        Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, Nothing)
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A Nothing repository is refused - description was [" & _
                    errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingTempRepos", Err.Number, Err.Description
End Sub

'@sub-title The three bindings are closed once Create has returned.
'@details
'Swapping the target workbook part way through a transfer would scatter
'components across two files.
'@TestMethod("CodeTransfer")
Public Sub TestTheBindingsAreRefusedAfterCreate()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBindingsAreRefusedAfterCreate"
    If Not FixtureReady("TestTheBindingsAreRefusedAfterCreate") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    On Error Resume Next
        Set sut.TargetWkb = SourceWkb
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), errNumber, _
                    "A sealed instance refuses a second target workbook - " & _
                    "description was [" & errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBindingsAreRefusedAfterCreate", Err.Number, Err.Description
End Sub


'@section Transfer tests
'===============================================================================

'@sub-title A standard module arrives in the target carrying its code.
'@TestMethod("CodeTransfer")
Public Sub TestTransferModuleCarriesTheCodeAcross()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferModuleCarriesTheCodeAcross"
    If Not FixtureReady("TestTransferModuleCarriesTheCodeAcross") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.TransferModule PROBE_MODULE

    Assert.AreEqual CLng(1), ComponentCount(TargetWkb, PROBE_MODULE), _
                    "The target holds the module once"
    Assert.IsTrue InStr(1, ComponentCode(TargetWkb, PROBE_MODULE), SOURCE_MARK, _
                        vbTextCompare) > 0, _
                  "The module in the target carries the source code"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferModuleCarriesTheCodeAcross", Err.Number, Err.Description
End Sub

'@sub-title A class module arrives in the target carrying its code.
'@TestMethod("CodeTransfer")
Public Sub TestTransferClassCarriesTheCodeAcross()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferClassCarriesTheCodeAcross"
    If Not FixtureReady("TestTransferClassCarriesTheCodeAcross") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.TransferClass PROBE_CLASS

    Assert.AreEqual CLng(1), ComponentCount(TargetWkb, PROBE_CLASS), _
                    "The target holds the class once"
    Assert.IsTrue InStr(1, ComponentCode(TargetWkb, PROBE_CLASS), SOURCE_MARK, _
                        vbTextCompare) > 0, _
                  "The class in the target carries the source code"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferClassCarriesTheCodeAcross", Err.Number, Err.Description
End Sub

'@sub-title A class the target already holds is replaced by the source copy.
'@details
'VBComponents.Import stays quiet on a name clash and leaves two copies under two
'names, which stops the project compiling. This is the template branch's normal
'case.
'@TestMethod("CodeTransfer")
Public Sub TestTransferClassReplacesTheCopyTheTargetHolds()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferClassReplacesTheCopyTheTargetHolds"
    If Not FixtureReady("TestTransferClassReplacesTheCopyTheTargetHolds") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim targetText As String

    AddComponent TargetWkb, PROBE_CLASS, VB_CLASS_MODULE, TARGET_LINE
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.TransferClass PROBE_CLASS
    targetText = ComponentCode(TargetWkb, PROBE_CLASS)

    Assert.AreEqual CLng(1), ComponentCount(TargetWkb, PROBE_CLASS), _
                    "The target still holds the class once"
    Assert.AreEqual CLng(0), ComponentCount(TargetWkb, PROBE_CLASS & "1"), _
                    "No second copy is left under a numbered name"
    Assert.IsTrue InStr(1, targetText, SOURCE_MARK, vbTextCompare) > 0, _
                  "The class that survives is the one from the source"
    Assert.IsTrue InStr(1, targetText, TARGET_MARK, vbTextCompare) = 0, _
                  "The code that was in the target is gone"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferClassReplacesTheCopyTheTargetHolds", Err.Number, Err.Description
End Sub

'@sub-title Replacing a component is filed so the generation report names it.
'@TestMethod("CodeTransfer")
Public Sub TestReplacingAComponentIsFiledInTheCheckings()
    CustomTestSetTitles Assert, TESTMODULE, "TestReplacingAComponentIsFiledInTheCheckings"
    If Not FixtureReady("TestReplacingAComponentIsFiledInTheCheckings") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim quiet As CodeTransfer

    AddComponent TargetWkb, PROBE_CLASS, VB_CLASS_MODULE, TARGET_LINE
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)
    Set quiet = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.TransferClass PROBE_CLASS
    quiet.TransferModule PROBE_MODULE

    Assert.IsTrue sut.HasCheckings, _
                  "A replaced component is filed"
    Assert.IsTrue Not sut.CheckingValues Is Nothing, _
                  "The filed entries are readable"
    Assert.IsFalse quiet.HasCheckings, _
                   "A transfer into a target without the component files nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestReplacingAComponentIsFiledInTheCheckings", Err.Number, Err.Description
End Sub

'@sub-title A component the source lacks is refused, and no file is left behind.
'@TestMethod("CodeTransfer")
Public Sub TestTransferClassRefusesAComponentTheSourceLacks()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferClassRefusesAComponentTheSourceLacks"
    If Not FixtureReady("TestTransferClassRefusesAComponentTheSourceLacks") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    On Error Resume Next
        sut.TransferClass PROBE_ABSENT
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A component the source lacks is refused - description was [" & _
                    errDescription & "]"
    Assert.AreEqual CLng(0), FileCountIn(TempRepos.RootPath), _
                    "The repository holds no file after the refusal"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferClassRefusesAComponentTheSourceLacks", Err.Number, Err.Description
End Sub

'@sub-title A finished transfer leaves the repository empty.
'@TestMethod("CodeTransfer")
Public Sub TestATransferLeavesNoTemporaryFile()
    CustomTestSetTitles Assert, TESTMODULE, "TestATransferLeavesNoTemporaryFile"
    If Not FixtureReady("TestATransferLeavesNoTemporaryFile") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.TransferClass PROBE_CLASS
    sut.TransferModule PROBE_MODULE

    Assert.AreEqual CLng(0), FileCountIn(TempRepos.RootPath), _
                    "Both export files are gone once the transfers are done"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestATransferLeavesNoTemporaryFile", Err.Number, Err.Description
End Sub


'@section Code writing tests
'===============================================================================

'@sub-title Code reaches the ThisWorkbook module of a workbook that has none.
'@details
'A workbook from Workbooks.Add carries an empty ThisWorkbook module, and
'DeleteLines 1, 0 on an empty module is the case the line-count guard exists
'for. This is the non-template branch of a generation.
'@TestMethod("CodeTransfer")
Public Sub TestTransferWorkbookCodeWritesIntoAnEmptyModule()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferWorkbookCodeWritesIntoAnEmptyModule"
    If Not FixtureReady("TestTransferWorkbookCodeWritesIntoAnEmptyModule") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    On Error Resume Next
        sut.TransferWorkbookCode PROBE_MODULE
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(0), errNumber, _
                    "An empty ThisWorkbook module takes the write - description was [" & _
                    errDescription & "]"
    Assert.IsTrue InStr(1, ComponentCode(TargetWkb, TargetWkb.CodeName), SOURCE_MARK, _
                        vbTextCompare) > 0, _
                  "The workbook module carries the source code"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferWorkbookCodeWritesIntoAnEmptyModule", Err.Number, Err.Description
End Sub

'@sub-title Code reaches the code module of a named worksheet.
'@TestMethod("CodeTransfer")
Public Sub TestTransferWorksheetCodeWritesIntoTheSheetModule()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferWorksheetCodeWritesIntoTheSheetModule"
    If Not FixtureReady("TestTransferWorksheetCodeWritesIntoTheSheetModule") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim sheetName As String

    sheetName = TargetWkb.Worksheets(1).Name
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.TransferWorksheetCode PROBE_MODULE, sheetName

    Assert.IsTrue InStr(1, ComponentCode(TargetWkb, TargetWkb.Worksheets(1).CodeName), _
                        SOURCE_MARK, vbTextCompare) > 0, _
                  "The worksheet module carries the source code"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferWorksheetCodeWritesIntoTheSheetModule", Err.Number, Err.Description
End Sub

'@sub-title A worksheet name the target does not carry is refused.
'@TestMethod("CodeTransfer")
Public Sub TestTransferWorksheetCodeRefusesAMissingSheet()
    CustomTestSetTitles Assert, TESTMODULE, "TestTransferWorksheetCodeRefusesAMissingSheet"
    If Not FixtureReady("TestTransferWorksheetCodeRefusesAMissingSheet") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    On Error Resume Next
        sut.TransferWorksheetCode PROBE_MODULE, "OBTSheetThatIsAbsent"
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A worksheet the target lacks is refused - description was [" & _
                    errDescription & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTransferWorksheetCodeRefusesAMissingSheet", Err.Number, Err.Description
End Sub

'@sub-title Copying text into a module both workbooks hold replaces the target text.
'@TestMethod("CodeTransfer")
Public Sub TestCopyModuleTextReplacesTheTargetText()
    CustomTestSetTitles Assert, TESTMODULE, "TestCopyModuleTextReplacesTheTargetText"
    If Not FixtureReady("TestCopyModuleTextReplacesTheTargetText") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim targetText As String

    AddComponent TargetWkb, PROBE_MODULE, VB_STANDARD_MODULE, TARGET_LINE
    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    sut.CopyModuleText PROBE_MODULE
    targetText = ComponentCode(TargetWkb, PROBE_MODULE)

    Assert.IsTrue InStr(1, targetText, SOURCE_MARK, vbTextCompare) > 0, _
                  "The target module carries the source code"
    Assert.IsTrue InStr(1, targetText, TARGET_MARK, vbTextCompare) = 0, _
                  "The code that was in the target is gone"
    Assert.AreEqual CLng(1), ComponentCount(TargetWkb, PROBE_MODULE), _
                    "The target holds the module once"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCopyModuleTextReplacesTheTargetText", Err.Number, Err.Description
End Sub

'@sub-title Copying text needs the module in the target, and says so when it is absent.
'@TestMethod("CodeTransfer")
Public Sub TestCopyModuleTextRefusesATargetWithoutTheModule()
    CustomTestSetTitles Assert, TESTMODULE, "TestCopyModuleTextRefusesATargetWithoutTheModule"
    If Not FixtureReady("TestCopyModuleTextRefusesATargetWithoutTheModule") Then Exit Sub
    On Error GoTo TestFail

    Dim sut As CodeTransfer
    Dim errNumber As Long
    Dim errDescription As String

    Set sut = CodeTransfer.Create(SourceWkb, TargetWkb, TempRepos)

    On Error Resume Next
        sut.CopyModuleText PROBE_MODULE
        errNumber = Err.Number
        errDescription = Err.Description
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "A target without the module is refused - description was [" & _
                    errDescription & "]"
    Assert.AreEqual CLng(0), ComponentCount(TargetWkb, PROBE_MODULE), _
                    "The refusal creates no component"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCopyModuleTextRefusesATargetWithoutTheModule", Err.Number, Err.Description
End Sub
