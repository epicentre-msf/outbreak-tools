Attribute VB_Name = "TestLinelistSaveWorkflow"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Navigator As ILinelistSheetNavigator
Private Lifecycle As ILinelistLifecycleManager
Private TempServiceStub As TemporaryReposStub
Private SaveWorkflow As ILinelistSaveWorkflow
Private Context As ILinelistPreparationContext
Private AccessorStub As LinelistRecordingAccessorStub
Private SpecsStub As LinelistSpecsWorkbookStub
Private MainStub As LinelistMainStub
Private PasswordStub As LinelistPasswordStub
Private DictionaryStub As DictionaryMinimalStub
Private TranslationStub As LinelistTranslationCounterStub
Private ScopeStub As ApplicationStateStub
Private WorkbookRef As Workbook
Private SavedOutputPath As String


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistSaveWorkflow"
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
    Dim formatter As ILinelistSheetNameFormatter
    Set formatter = LinelistSheetNameFormatter.Create

    Set WorkbookRef = TestHelpers.NewWorkbook
    WorkbookRef.Worksheets(1).Name = "Administration"
    If WorkbookRef.Worksheets.Count < 2 Then WorkbookRef.Worksheets.Add
    WorkbookRef.Worksheets(2).Name = "Instructions"

    Set DictionaryStub = New DictionaryMinimalStub

    Set SpecsStub = New LinelistSpecsWorkbookStub
    SpecsStub.Initialise DictionaryStub, WorkbookRef

    Set MainStub = New LinelistMainStub
    MainStub.Initialise

    Dim tempDirectory As String
    tempDirectory = Application.DefaultFilePath
    If LenB(Trim$(tempDirectory)) = 0 Then tempDirectory = Environ$("TMPDIR")
    If LenB(Trim$(tempDirectory)) = 0 Then tempDirectory = Environ$("TEMP")
    If LenB(Trim$(tempDirectory)) = 0 Then tempDirectory = CurDir$

    Dim uniqueSuffix As String
    uniqueSuffix = Format$(Now, "yyyymmdd_hhnnss")

    MainStub.SetValue "lldir", tempDirectory
    MainStub.SetValue "llname", "linelist_tests_" & uniqueSuffix
    MainStub.SetValue "llpassword", "secret"

    SavedOutputPath = tempDirectory & Application.PathSeparator & "linelist_tests_" & uniqueSuffix & ".xlsb"

    Set PasswordStub = New LinelistPasswordStub

    Set TranslationStub = New LinelistTranslationCounterStub
    TranslationStub.Initialise
    TranslationStub.SetValue "LLSHEET_Admin", "Administration"
    TranslationStub.SetValue "INSTSHEETNAME", "Instructions"

    SpecsStub.SetMainObject MainStub
    SpecsStub.SetPassword PasswordStub
    SpecsStub.SetTranslations TranslationStub

    Set AccessorStub = New LinelistRecordingAccessorStub
    AccessorStub.Initialise DictionaryStub, SpecsStub, WorkbookRef

    Set TempServiceStub = New TemporaryReposStub

    Set ScopeStub = New ApplicationStateStub
    Set ScopeStub.ApplicationObject = Application

    Set Navigator = LinelistSheetNavigator.Create(WorkbookRef, TranslationStub, formatter)
    Set Lifecycle = LinelistLifecycleManager.Create(AccessorStub, TempServiceStub, ScopeStub)

    Set SaveWorkflow = LinelistSaveWorkflow.Create(Navigator, Lifecycle)
    Set Context = LinelistPreparationContext.Create(AccessorStub, SpecsStub, DictionaryStub, formatter, ScopeStub)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If LenB(SavedOutputPath) > 0 Then
        On Error Resume Next
            Kill SavedOutputPath
        On Error GoTo 0
    End If

    If Not WorkbookRef Is Nothing Then
        BusyApp
        On Error Resume Next
            WorkbookRef.Close saveChanges:=False
        On Error GoTo 0
        Set WorkbookRef = Nothing
    End If

    Set SaveWorkflow = Nothing
    Set Navigator = Nothing
    Set Lifecycle = Nothing
    Set TempServiceStub = Nothing
    Set Context = Nothing
    Set AccessorStub = Nothing
    Set SpecsStub = Nothing
    Set MainStub = Nothing
    Set PasswordStub = Nothing
    Set DictionaryStub = Nothing
    Set TranslationStub = Nothing
    Set ScopeStub = Nothing
End Sub


'@section Tests
'===============================================================================
'@TestMethod("LinelistSaveWorkflow")
Public Sub TestSaveWorkflowIntegration()
    CustomTestSetTitles Assert, "LinelistSaveWorkflow", "SaveWorkflowIntegration"

    SaveWorkflow.Save Context
    Set WorkbookRef = Nothing 'Lifecycle closes the workbook

    Assert.AreEqual 2, TranslationStub.LookupCount, "Translations should be resolved for admin and instruction sheets"
    Assert.AreEqual "Instructions", Navigator.LastActivatedSheetName, "Instruction sheet should be activated last"
    Assert.AreEqual 1, PasswordStub.ProtectCount, "Workbook should be protected before saving"
    Assert.AreEqual 1, TempServiceStub.DeleteAllCount, "Temp files should be cleared once"
    Assert.AreEqual 1, TempServiceStub.ResetCount, "Temp directory should be reset once"
    Assert.AreEqual 1, AccessorStub.ClearCount, "Workbook reference should be cleared after dispose"
    Assert.IsTrue Dir$(SavedOutputPath) <> vbNullString, "Saved workbook should exist on disk"
End Sub

'@TestMethod("LinelistSaveWorkflow")
Public Sub TestSaveRaisesWhenWorkbookMissing()
    CustomTestSetTitles Assert, "LinelistSaveWorkflow", "SaveRaisesWhenWorkbookMissing"

    AccessorStub.Initialise DictionaryStub, SpecsStub, Nothing

    On Error GoTo ExpectError
        SaveWorkflow.Save Context
        Assert.Fail "Missing workbook should raise"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Missing workbook should raise ObjectNotInitialized"
    Err.Clear
End Sub

