Attribute VB_Name = "TestLinelistCodeTransferStep"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private ServiceStub As LinelistCodeTransferServiceStub
Private PlanStub As LinelistCodeTransferPlanStub
Private SpecsStub As LinelistSpecsWorkbookStub
Private LinelistStub As LinelistWorkbookAccessorStub
Private ScopeStub As ApplicationStateStub
Private DictionaryStub As DictionaryMinimalStub
Private SourceWorkbook As Workbook
Private TargetWorkbook As Workbook


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistCodeTransferStep"
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
    Set ServiceStub = New LinelistCodeTransferServiceStub
    ServiceStub.Initialise

    Set PlanStub = New LinelistCodeTransferPlanStub
    PlanStub.Initialise

    Set DictionaryStub = New DictionaryMinimalStub

    Set SpecsStub = New LinelistSpecsWorkbookStub

    Set LinelistStub = New LinelistWorkbookAccessorStub

    Set ScopeStub = New ApplicationStateStub
    Set ScopeStub.ApplicationObject = Application

    Set SourceWorkbook = TestHelpers.NewWorkbook
    Set TargetWorkbook = TestHelpers.NewWorkbook

    SpecsStub.Initialise DictionaryStub, SourceWorkbook
    LinelistStub.Initialise DictionaryStub, SpecsStub, TargetWorkbook
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

    Set ServiceStub = Nothing
    Set PlanStub = Nothing
    Set SpecsStub = Nothing
    Set LinelistStub = Nothing
    Set ScopeStub = Nothing
    Set DictionaryStub = Nothing
End Sub


'@section Helper builders
'===============================================================================
Private Function BuildContext() As ILinelistPreparationContext
    Dim formatter As ILinelistSheetNameFormatter
    Set formatter = LinelistSheetNameFormatter.Create

    Set BuildContext = LinelistPreparationContext.Create(LinelistStub, SpecsStub, _
                                                         DictionaryStub, formatter, ScopeStub)
End Function

Private Function CreateStep(Optional ByVal plan As ILinelistCodeTransferPlan) As ILinelistPreparationStep
    Dim planToUse As ILinelistCodeTransferPlan
    If plan Is Nothing Then
        Set planToUse = PlanStub
    Else
        Set planToUse = plan
    End If

    Set CreateStep = LinelistCodeTransferStep.Create(ServiceStub, planToUse)
End Function

Private Sub ConfigurePlan()
    PlanStub.SetClassModules "ClassA", "ClassB"
    PlanStub.SetStandardModules "ModuleA"
    PlanStub.SetForms "FormA", "FormB"
    PlanStub.WorkbookModuleName = "WorkbookModule"
End Sub


'@section Tests
'===============================================================================
'@TestMethod("LinelistCodeTransferStep")
Public Sub TestExecuteTransfersAllArtifacts()
    CustomTestSetTitles Assert, "LinelistCodeTransferStep", "ExecuteTransfersAllArtifacts"

    ConfigurePlan

    Dim stepObject As ILinelistPreparationStep
    Set stepObject = CreateStep()

    Dim context As ILinelistPreparationContext
    Set context = BuildContext()

    stepObject.Execute context

    Assert.AreEqual 3, ServiceStub.ComponentCalls.Length, _
                     "Service should receive two class entries and one module entry"
    Assert.AreEqual "ClassA|" & CStr(codeScopeClass), CStr(ServiceStub.ComponentCalls.Item(1)), _
                     "First call should target class module ClassA"
    Assert.AreEqual "ModuleA|" & CStr(codeScopeModule), CStr(ServiceStub.ComponentCalls.Item(3)), _
                     "Last call should target standard module ModuleA"

    Assert.AreEqual 2, ServiceStub.FormCalls.Length, "Two forms should be requested"
    Assert.AreEqual "FormB", CStr(ServiceStub.FormCalls.Item(2)), "Second form should be FormB"

    Assert.AreEqual 1, ServiceStub.WorkbookModuleCalls.Length, "Workbook module should be transferred"
    Assert.AreEqual "WorkbookModule", CStr(ServiceStub.WorkbookModuleCalls.Item(1)), _
                     "Workbook module name should match plan"
End Sub

'@TestMethod("LinelistCodeTransferStep")
Public Sub TestExecuteSkipsWorkbookModuleWhenNameEmpty()
    CustomTestSetTitles Assert, "LinelistCodeTransferStep", "ExecuteSkipsWorkbookModuleWhenNameEmpty"

    PlanStub.SetClassModules "OnlyClass"
    PlanStub.SetStandardModules
    PlanStub.SetForms
    PlanStub.WorkbookModuleName = vbNullString

    Dim stepObject As ILinelistPreparationStep
    Set stepObject = CreateStep()

    Dim context As ILinelistPreparationContext
    Set context = BuildContext()

    stepObject.Execute context

    Assert.AreEqual 1, ServiceStub.ComponentCalls.Length, "One class component should be transferred"
    Assert.AreEqual 0, ServiceStub.WorkbookModuleCalls.Length, "Workbook module transfer should be skipped"
End Sub

'@TestMethod("LinelistCodeTransferStep")
Public Sub TestExecuteRaisesWhenLinelistMissingAccessor()
    CustomTestSetTitles Assert, "LinelistCodeTransferStep", "ExecuteRaisesWhenLinelistMissingAccessor"

    ConfigurePlan

    Dim stepObject As ILinelistPreparationStep
    Set stepObject = CreateStep()

    Dim context As ILinelistPreparationContext
    Dim formatter As ILinelistSheetNameFormatter
    Set formatter = LinelistSheetNameFormatter.Create

    Dim plainLinelist As LinelistDictionaryStub
    Set plainLinelist = New LinelistDictionaryStub
    plainLinelist.Initialise DictionaryStub

    Set context = LinelistPreparationContext.Create(plainLinelist, SpecsStub, _
                                                    DictionaryStub, formatter, ScopeStub)

    On Error GoTo ExpectError
        stepObject.Execute context
        Assert.Fail "Executing without ILinelistWorkbookAccessor should raise"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidOperation, Err.Number, _
                     "Missing accessor should raise InvalidOperation"
    Err.Clear
End Sub

