Attribute VB_Name = "TestLinelistPreparationPipeline"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private CallLog As BetterArray


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistPreparationPipeline"
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
    Set CallLog = New BetterArray
    CallLog.LowerBound = 1
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Helper builders
'===============================================================================
Private Function CreateSpecs(Optional ByVal withDictionary As Boolean = True) As ILinelistSpecs
    Dim specs As TableSpecsLinelistStub
    Dim dictionary As DictionaryMinimalStub

    Set specs = New TableSpecsLinelistStub
    If withDictionary Then
        Set dictionary = New DictionaryMinimalStub
        specs.SetDictionary dictionary
    End If

    Set CreateSpecs = specs
End Function

Private Function CreateLinelist(ByVal dictionary As ILLdictionary) As ILinelist
    Dim linelistStub As LinelistDictionaryStub
    Set linelistStub = New LinelistDictionaryStub
    linelistStub.Initialise dictionary
    Set CreateLinelist = linelistStub
End Function

Private Function CreateStep(ByVal label As String, Optional ByVal shouldFail As Boolean = False) As ILinelistPreparationStep
    Dim stub As LinelistPreparationStepStub
    Set stub = New LinelistPreparationStepStub
    stub.Initialise label, CallLog
    stub.ShouldFail = shouldFail
    Set CreateStep = stub
End Function

Private Function CreateScopeStub() As ILinelistApplicationStateScope
    Dim scopeStub As LinelistApplicationStateScopeStub
    Set scopeStub = New LinelistApplicationStateScopeStub
    Set scopeStub.ApplicationObject = Application
    Set CreateScopeStub = scopeStub
End Function

Private Function BuildPipeline(Optional ByVal failingStage As LinelistPreparationStage = 0) As ILinelistPreparationPipeline
    Dim specs As ILinelistSpecs
    Dim dictionary As ILLdictionary
    Dim linelist As ILinelist
    Dim scope As ILinelistApplicationStateScope
    Dim tempStep As ILinelistPreparationStep
    Dim exportStep As ILinelistPreparationStep
    Dim analysisStep As ILinelistPreparationStep
    Dim codeStep As ILinelistPreparationStep
    Dim formStep As ILinelistPreparationStep

    Set specs = CreateSpecs(True)
    Set dictionary = specs.Dictionary
    Set linelist = CreateLinelist(dictionary)
    Set scope = CreateScopeStub()

    Set tempStep = CreateStep("CreateTemporarySheets", failingStage = LinelistPreparationPipeline.stageCreateTemporarySheets)
    Set exportStep = CreateStep("ExportSpecifications", failingStage = LinelistPreparationPipeline.stageExportSpecifications)
    Set analysisStep = CreateStep("CreateAnalysisSheets", failingStage = LinelistPreparationPipeline.stageCreateAnalysisSheets)
    Set codeStep = CreateStep("TransferCodeModules", failingStage = LinelistPreparationPipeline.stageTransferCodeModules)
    Set formStep = CreateStep("TransferForms", failingStage = LinelistPreparationPipeline.stageTransferForms)

    Set BuildPipeline = LinelistPreparationPipeline.Create(linelist, specs, _
                                                           tempStep, exportStep, analysisStep, codeStep, formStep, _
                                                           LinelistSheetNameFormatter.Create, scope)
End Function


'@section Tests
'===============================================================================
'@TestMethod("LinelistPreparationPipeline")
Public Sub TestPrepareRunsStepsInOrder()
    CustomTestSetTitles Assert, "LinelistPreparationPipeline", "PrepareRunsStepsInOrder"

    Dim pipeline As ILinelistPreparationPipeline
    Set pipeline = BuildPipeline()

    pipeline.Prepare

    Assert.AreEqual 5, CallLog.Length, "Pipeline should execute five stages"
    Assert.AreEqual "CreateTemporarySheets", CStr(CallLog.Item(1)), "Temporary sheets stage should run first"
    Assert.AreEqual "ExportSpecifications", CStr(CallLog.Item(2)), "Specification export should follow temporary sheets"
    Assert.AreEqual "CreateAnalysisSheets", CStr(CallLog.Item(3)), "Analysis sheet creation should run third"
    Assert.AreEqual "TransferCodeModules", CStr(CallLog.Item(4)), "Code transfer should run fourth"
    Assert.AreEqual "TransferForms", CStr(CallLog.Item(5)), "Form transfer should run last"
End Sub

'@TestMethod("LinelistPreparationPipeline")
Public Sub TestPrepareRestoresScopeWhenStepFails()
    CustomTestSetTitles Assert, "LinelistPreparationPipeline", "PrepareRestoresScopeWhenStepFails"

    Dim pipeline As ILinelistPreparationPipeline
    Dim scopeStub As LinelistApplicationStateScopeStub

    Set scopeStub = CreateScopeStub()

    Dim specs As ILinelistSpecs
    Dim linelist As ILinelist
    Dim dictionary As ILLdictionary
    Dim tempStep As ILinelistPreparationStep
    Dim exportStep As ILinelistPreparationStep
    Dim analysisStep As ILinelistPreparationStep
    Dim codeStep As ILinelistPreparationStep
    Dim formStep As ILinelistPreparationStep

    Set specs = CreateSpecs(True)
    Set dictionary = specs.Dictionary
    Set linelist = CreateLinelist(dictionary)
    Set tempStep = CreateStep("CreateTemporarySheets")
    Set exportStep = CreateStep("ExportSpecifications")
    Set analysisStep = CreateStep("CreateAnalysisSheets")
    Set codeStep = CreateStep("TransferCodeModules", shouldFail:=True)
    Set formStep = CreateStep("TransferForms")

    Dim scopeInterface As ILinelistApplicationStateScope
    Set scopeInterface = scopeStub

    Set pipeline = LinelistPreparationPipeline.Create(linelist, specs, _
                                                      tempStep, exportStep, analysisStep, codeStep, formStep, _
                                                      LinelistSheetNameFormatter.Create, scopeInterface)

    On Error GoTo ExpectError
        pipeline.Prepare
        Assert.Fail "Prepare should bubble an error when a stage fails"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ErrorUnexpectedState, Err.Number, "Pipeline should translate failures to ProjectError.ErrorUnexpectedState"
    Err.Clear

    Assert.AreEqual 1, scopeStub.ApplyCount, "ApplyBusyState should be invoked once"
    Assert.AreEqual 1, scopeStub.RestoreCount, "Restore must be called once even on failure"
End Sub

'@TestMethod("LinelistPreparationPipeline")
Public Sub TestCreateRaisesWhenDictionaryMissing()
    CustomTestSetTitles Assert, "LinelistPreparationPipeline", "CreateRaisesWhenDictionaryMissing"

    Dim specs As ILinelistSpecs
    Dim linelist As ILinelist
    Dim dictionary As ILLdictionary
    Dim tempStep As ILinelistPreparationStep

    Set specs = CreateSpecs(False)
    Set dictionary = Nothing
    Set linelist = CreateLinelist(New DictionaryMinimalStub)
    Set tempStep = CreateStep("CreateTemporarySheets")

    On Error GoTo ExpectError
        Call LinelistPreparationPipeline.Create(linelist, specs, _
                                                tempStep, tempStep, tempStep, tempStep, tempStep, _
                                                LinelistSheetNameFormatter.Create, CreateScopeStub())
        Assert.Fail "Pipeline creation should raise when specifications expose no dictionary"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Missing dictionary should raise ObjectNotInitialized"
    Err.Clear
End Sub
