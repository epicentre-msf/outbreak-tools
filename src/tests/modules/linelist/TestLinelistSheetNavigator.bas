Attribute VB_Name = "TestLinelistSheetNavigator"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Navigator As ILinelistSheetNavigator
Private Formatter As ILinelistSheetNameFormatter
Private TranslationStub As LinelistTranslationCounterStub
Private WorkbookRef As Workbook


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistSheetNavigator"
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
    Set Formatter = LinelistSheetNameFormatter.Create

    Set TranslationStub = New LinelistTranslationCounterStub
    TranslationStub.Initialise
    TranslationStub.SetValue "LLSHEET_Admin", "Administration"
    TranslationStub.SetValue "INSTSHEETNAME", "Instructions"

    Set WorkbookRef = TestHelpers.NewWorkbook
    PrepareWorkbookSheets

    Set Navigator = LinelistSheetNavigator.Create(WorkbookRef, TranslationStub, Formatter)
End Sub

Private Sub PrepareWorkbookSheets()
    Dim adminSheet As Worksheet
    Dim instructionSheet As Worksheet

    Set adminSheet = WorkbookRef.Worksheets(1)
    adminSheet.Name = "Administration"

    If WorkbookRef.Worksheets.Count < 2 Then
        WorkbookRef.Worksheets.Add
    End If

    Set instructionSheet = WorkbookRef.Worksheets(2)
    instructionSheet.Name = "Instructions"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not WorkbookRef Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook WorkbookRef
        Set WorkbookRef = Nothing
    End If

    Set Navigator = Nothing
    Set Formatter = Nothing
    Set TranslationStub = Nothing
End Sub


'@section Tests
'===============================================================================
'@TestMethod("LinelistSheetNavigator")
Public Sub TestActivateAdminAndInstructionSheets()
    CustomTestSetTitles Assert, "LinelistSheetNavigator", "ActivateAdminAndInstructionSheets"

    Navigator.ActivateAdminSheet
    Navigator.ActivateInstructionSheet

    Dim instructionName As String
    instructionName = WorkbookRef.ActiveSheet.Name

    Navigator.ActivateAdminSheet 'ensure cache reused

    Assert.AreEqual "Administration", Navigator.LastActivatedSheetName, _
                     "Last activated sheet should reflect most recent activation"
    Assert.AreEqual 2, Navigator.TranslationCacheHits, _
                     "Translations should only be resolved once per key"
    Assert.AreEqual 2, TranslationStub.LookupCount, _
                     "Translation stub should only be queried twice"

    Assert.AreEqual "Instructions", instructionName, _
                     "Instruction activation should select the instruction worksheet"
End Sub

'@TestMethod("LinelistSheetNavigator")
Public Sub TestActivateSheetRespectsScope()
    CustomTestSetTitles Assert, "LinelistSheetNavigator", "ActivateSheetRespectsScope"

    Dim scopedSheet As Worksheet
    Set scopedSheet = WorkbookRef.Worksheets.Add
    scopedSheet.Name = Formatter.FormatSheetName("Report", sheetScopePrint)

    Dim activated As Worksheet
    Set activated = Navigator.ActivateSheet("Report", sheetScopePrint)

    Assert.AreEqual scopedSheet.Name, activated.Name, "Activation should return the scoped worksheet"
    Assert.IsTrue Navigator.SheetExists("Report", sheetScopePrint), "SheetExists should report sheet presence"
    Assert.AreEqual scopedSheet.Name, Navigator.LastActivatedSheetName, _
                     "Navigator should remember last scoped sheet name"
End Sub

'@TestMethod("LinelistSheetNavigator")
Public Sub TestResolveSheetRaisesWhenMissing()
    CustomTestSetTitles Assert, "LinelistSheetNavigator", "ResolveSheetRaisesWhenMissing"

    On Error GoTo ExpectMissing
        Navigator.ResolveSheet "DoesNotExist"
        Assert.Fail "Resolving a missing sheet should raise"
        Exit Sub
ExpectMissing:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing sheet should raise ElementNotFound"
    Err.Clear
End Sub
