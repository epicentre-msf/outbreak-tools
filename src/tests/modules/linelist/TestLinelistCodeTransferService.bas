Attribute VB_Name = "TestLinelistCodeTransferService"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private FixtureWorkbook As Workbook
Private StrategyStub As LinelistCodeTransferStrategyStub


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistCodeTransferService"
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
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set StrategyStub = New LinelistCodeTransferStrategyStub
    StrategyStub.Initialise "StubStrategy"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not FixtureWorkbook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook FixtureWorkbook
        Set FixtureWorkbook = Nothing
    End If

    Set StrategyStub = Nothing
End Sub


'@section Helper builders
'===============================================================================
Private Function CreateService() As ILinelistCodeTransferService
    Set CreateService = LinelistCodeTransferService.Create(StrategyStub)
End Function


'@section Tests
'===============================================================================
'@TestMethod("LinelistCodeTransferService")
Public Sub TestTransferComponentDelegatesToStrategy()
    CustomTestSetTitles Assert, "LinelistCodeTransferService", "TransferComponentDelegatesToStrategy"

    Dim service As ILinelistCodeTransferService
    Set service = CreateService()

    service.TransferComponent "ModuleA", codeScopeModule, FixtureWorkbook, FixtureWorkbook

    Assert.AreEqual 1, StrategyStub.ComponentLog.Length, "Strategy should receive one component transfer call"
    Assert.AreEqual "ModuleA|" & CStr(codeScopeModule), CStr(StrategyStub.ComponentLog.Item(1)), _
                     "Strategy should record the component name and scope"
End Sub

'@TestMethod("LinelistCodeTransferService")
Public Sub TestTransferComponentsIteratesList()
    CustomTestSetTitles Assert, "LinelistCodeTransferService", "TransferComponentsIteratesList"

    Dim service As ILinelistCodeTransferService
    Dim moduleList As BetterArray

    Set service = CreateService()
    Set moduleList = New BetterArray
    moduleList.LowerBound = 1
    moduleList.Push "ModuleA", "ModuleB", "ModuleC"

    service.TransferComponents moduleList, codeScopeClass, FixtureWorkbook, FixtureWorkbook

    Assert.AreEqual 3, StrategyStub.ComponentLog.Length, "All modules should be forwarded to the strategy"
    Assert.AreEqual "ModuleA|" & CStr(codeScopeClass), CStr(StrategyStub.ComponentLog.Item(1)), _
                     "First module should be ModuleA"
    Assert.AreEqual "ModuleC|" & CStr(codeScopeClass), CStr(StrategyStub.ComponentLog.Item(3)), _
                     "Last module should be ModuleC"
End Sub

'@TestMethod("LinelistCodeTransferService")
Public Sub TestTransferComponentRejectsEmptyName()
    CustomTestSetTitles Assert, "LinelistCodeTransferService", "TransferComponentRejectsEmptyName"

    Dim service As ILinelistCodeTransferService
    Set service = CreateService()

    On Error GoTo ExpectError
        service.TransferComponent vbNullString, codeScopeModule, FixtureWorkbook, FixtureWorkbook
        Assert.Fail "Empty component name should raise"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.InvalidArgument, Err.Number, _
                     "Empty component name should raise InvalidArgument"
    Err.Clear
End Sub

'@TestMethod("LinelistCodeTransferService")
Public Sub TestTransferComponentRejectsMissingWorkbook()
    CustomTestSetTitles Assert, "LinelistCodeTransferService", "TransferComponentRejectsMissingWorkbook"

    Dim service As ILinelistCodeTransferService
    Set service = CreateService()

    On Error GoTo ExpectError
        service.TransferComponent "ModuleA", codeScopeModule, Nothing, FixtureWorkbook
        Assert.Fail "Missing source workbook should raise"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, _
                     "Missing workbook should raise ObjectNotInitialized"
    Err.Clear
End Sub

'@TestMethod("LinelistCodeTransferService")
Public Sub TestStrategyNameReflectsUnderlyingStrategy()
    CustomTestSetTitles Assert, "LinelistCodeTransferService", "StrategyNameReflectsUnderlyingStrategy"

    Dim service As ILinelistCodeTransferService
    Set service = CreateService()

    Assert.AreEqual "StubStrategy", service.StrategyName, _
                     "StrategyName should expose the underlying strategy label"
End Sub

