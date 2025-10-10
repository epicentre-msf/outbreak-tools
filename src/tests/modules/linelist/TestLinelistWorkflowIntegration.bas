Attribute VB_Name = "TestLinelistWorkflowIntegration"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Pipeline As ILinelistPreparationPipeline
Private Context As ILinelistPreparationContext
Private LinelistStub As LinelistRecordingAccessorStub
Private SpecsStub As LinelistSpecsWorkbookStub
Private DictionaryStub As DictionaryMinimalStub
Private CodePlanStub As LinelistCodeTransferPlanStub
Private StrategyStub As LinelistCodeTransferStrategyStub
Private SourceWorkbook As Workbook
Private TargetWorkbook As Workbook


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistWorkflowIntegration"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If
    Set Assert = Nothing
    RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set DictionaryStub = New DictionaryMinimalStub

    Set SourceWorkbook = TestHelpers.NewWorkbook
    Set TargetWorkbook = TestHelpers.NewWorkbook

    Set SpecsStub = New LinelistSpecsWorkbookStub
    SpecsStub.Initialise DictionaryStub, SourceWorkbook

    Set LinelistStub = New LinelistRecordingAccessorStub
    LinelistStub.Initialise DictionaryStub, SpecsStub, TargetWorkbook

    Set StrategyStub = New LinelistCodeTransferStrategyStub
    StrategyStub.Initialise "StrategyStub"

    Set CodePlanStub = New LinelistCodeTransferPlanStub
    CodePlanStub.Initialise
    CodePlanStub.SetClassModules "ClassOne", "ClassTwo"
    CodePlanStub.SetStandardModules "ModuleOne"
    CodePlanStub.SetForms "FormOne"
    CodePlanStub.WorkbookModuleName = "WorkbookModule"

    Dim formatter As ILinelistSheetNameFormatter
    Set formatter = LinelistSheetNameFormatter.Create

    Dim scopeStub As ApplicationStateStub
    Set scopeStub = New ApplicationStateStub
    Set scopeStub.ApplicationObject = Application

    Set Context = LinelistPreparationContext.Create(LinelistStub, SpecsStub, DictionaryStub, formatter, scopeStub)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not SourceWorkbook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook SourceWorkbook
        Set SourceWorkbook = Nothing
    End If

    If Not TargetWorkbook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook TargetWorkbook
        Set TargetWorkbook = Nothing
    End If

    Set Pipeline = Nothing
    Set Context = Nothing
    Set LinelistStub = Nothing
    Set SpecsStub = Nothing
    Set DictionaryStub = Nothing
    Set CodePlanStub = Nothing
    Set StrategyStub = Nothing
End Sub


'@section Helper builders
'===============================================================================
Private Function BuildPipeline() As ILinelistPreparationPipeline
    Dim tempStep As LinelistSheetCreationStep
    Dim analysisStep As LinelistSheetCreationStep
    Dim noopStep As LinelistNoOpStep
    Dim codeService As ILinelistCodeTransferService
    Dim codeStep As ILinelistPreparationStep

    Set tempStep = New LinelistSheetCreationStep
    tempStep.Initialise "TempSheetA", "TempSheetB"

    Set analysisStep = New LinelistSheetCreationStep
    analysisStep.Initialise "AnalysisSheet"

    Set noopStep = New LinelistNoOpStep

    Set codeService = LinelistCodeTransferService.Create(StrategyStub)
    Set codeStep = LinelistCodeTransferStep.Create(codeService, CodePlanStub)

    Set BuildPipeline = LinelistPreparationPipeline.Create(LinelistStub, SpecsStub, _
                                                           tempStep, noopStep, analysisStep, _
                                                           codeStep, noopStep, _
                                                           LinelistSheetNameFormatter.Create)
End Function


'@section Tests
'===============================================================================
'@TestMethod("LinelistWorkflowIntegration")
Public Sub TestPrepareCoordinatesSheetsAndModuleTransfer()
    CustomTestSetTitles Assert, "LinelistWorkflowIntegration", "PrepareCoordinatesSheetsAndModuleTransfer"

    Set Pipeline = BuildPipeline()
    Pipeline.Prepare

    Dim createdSheets As BetterArray
    Set createdSheets = LinelistStub.CreatedSheets

    Assert.AreEqual 3, createdSheets.Length, "Pipeline should create three sheets"
    Assert.AreEqual "TempSheetA", CStr(createdSheets.Item(1)), "First sheet should be TempSheetA"
    Assert.AreEqual "TempSheetB", CStr(createdSheets.Item(2)), "Second sheet should be TempSheetB"
    Assert.AreEqual "AnalysisSheet", CStr(createdSheets.Item(3)), "Third sheet should be AnalysisSheet"

    Assert.AreEqual 2, StrategyStub.ComponentLog.Length, "Two class components should be transferred"
    Assert.AreEqual 1, StrategyStub.WorkbookModuleLog.Length, "Workbook module should be transferred"
    Assert.AreEqual 1, StrategyStub.FormLog.Length, "One form should be transferred"
End Sub

