Attribute VB_Name = "TestLinelistLifecycleManager"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Lifecycle As ILinelistLifecycleManager
Private TempServiceStub As LinelistTempFileServiceStub
Private ScopeStub As LinelistApplicationStateScopeStub
Private AccessorStub As LinelistWorkbookAccessorStub
Private DictionaryStub As DictionaryMinimalStub
Private SpecsStub As LinelistSpecsWorkbookStub
Private WorkbookRef As Workbook


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistLifecycleManager"
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
    Set TempServiceStub = New LinelistTempFileServiceStub
    Set ScopeStub = New LinelistApplicationStateScopeStub
    Set ScopeStub.ApplicationObject = Application

    Set DictionaryStub = New DictionaryMinimalStub
    Set SpecsStub = New LinelistSpecsWorkbookStub
    Set AccessorStub = New LinelistWorkbookAccessorStub

    Set WorkbookRef = TestHelpers.NewWorkbook
    SpecsStub.Initialise DictionaryStub, WorkbookRef
    AccessorStub.Initialise DictionaryStub, SpecsStub, WorkbookRef

    Set Lifecycle = LinelistLifecycleManager.Create(AccessorStub, TempServiceStub, ScopeStub)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    On Error Resume Next
        If Not WorkbookRef Is Nothing Then
            TestHelpers.DeleteWorkbook WorkbookRef
            Set WorkbookRef = Nothing
        End If
    On Error GoTo 0

    Set Lifecycle = Nothing
    Set TempServiceStub = Nothing
    Set ScopeStub = Nothing
    Set AccessorStub = Nothing
    Set DictionaryStub = Nothing
    Set SpecsStub = Nothing
End Sub


'@section Helper functions
'===============================================================================
Private Function WorkbookIsClosed(ByVal workbook As Workbook) As Boolean
    On Error GoTo WasClosed
        Dim nameValue As String
        nameValue = workbook.Name
        WorkbookIsClosed = False
        Exit Function
WasClosed:
    Err.Clear
    WorkbookIsClosed = True
End Function


'@section Tests
'===============================================================================
'@TestMethod("LinelistLifecycleManager")
Public Sub TestResetClosesWorkbookAndCleansTempFiles()
    CustomTestSetTitles Assert, "LinelistLifecycleManager", "ResetClosesWorkbookAndCleansTempFiles"

    Lifecycle.Reset False

    Assert.AreEqual 1, TempServiceStub.DeleteAllCount, "Reset should delete all temporary files"
    Assert.AreEqual 1, TempServiceStub.ResetCount, "Reset should rebuild temporary folder state"
    Assert.AreEqual 1, AccessorStub.ClearCount, "Reset should clear the cached workbook reference"
    Assert.AreEqual 1, ScopeStub.RefreshCount, "Reset should refresh the application state snapshot"

    Dim workbookAfter As Workbook
    Set workbookAfter = AccessorStub.OutputWorkbook
    Assert.IsTrue workbookAfter Is Nothing, "Workbook reference should be cleared"

    Assert.IsTrue WorkbookIsClosed(WorkbookRef), "Workbook should be closed after reset"
    Set WorkbookRef = Nothing
End Sub

'@TestMethod("LinelistLifecycleManager")
Public Sub TestDisposeClearsDependencies()
    CustomTestSetTitles Assert, "LinelistLifecycleManager", "DisposeClearsDependencies"

    Lifecycle.Dispose False
    Set WorkbookRef = Nothing

    On Error GoTo ExpectDisposed
        Lifecycle.Reset False
        Assert.Fail "Reset after dispose should raise"
        Exit Sub
ExpectDisposed:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Reset should fail once lifecycle manager is disposed"
    Err.Clear
End Sub

'@TestMethod("LinelistLifecycleManager")
Public Sub TestResetWithNoActiveWorkbook()
    CustomTestSetTitles Assert, "LinelistLifecycleManager", "ResetWithNoActiveWorkbook"

    Lifecycle.Reset False
    Set WorkbookRef = Nothing

    Assert.AreEqual 1, AccessorStub.ClearCount, "First reset should clear workbook once"

    Lifecycle.Reset False

    Assert.AreEqual 2, TempServiceStub.DeleteAllCount, "Each reset should purge temporary files"
    Assert.AreEqual 2, TempServiceStub.ResetCount, "Each reset should rebuild temporary folder"
    Assert.AreEqual 2, AccessorStub.ClearCount, "Clear should be invoked even when no workbook remains"
End Sub
