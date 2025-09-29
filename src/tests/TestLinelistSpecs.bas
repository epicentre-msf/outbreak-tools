Attribute VB_Name = "TestLinelistSpecs"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Smoke tests for the modernised linelist specifications facade")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private SpecsWorkbook As Workbook
Private Specs As ILinelistSpecs
Private MainStub As LinelistSpecsMainStub
Private DesignerStub As LinelistSpecsDesignerTranslationStub
Private MessageStub As LinelistSpecsTranslationStub

Private Const TEST_DESIGN_NAME As String = "UnitTestDesign"

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set SpecsWorkbook = TestHelpers.NewWorkbook()
    PrepareSpecificationWorkbook SpecsWorkbook
    Set Specs = LinelistSpecs.Create(SpecsWorkbook)
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Specs = Nothing
    Set MessageStub = Nothing
    Set DesignerStub = Nothing
    Set MainStub = Nothing
    TestHelpers.DeleteWorkbook SpecsWorkbook
    Set SpecsWorkbook = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Specs.ResetCaches

    Set MainStub = New LinelistSpecsMainStub
    MainStub.Initialise True
    MainStub.SetValue "setuppath", vbNullString
    MainStub.SetValue "geopath", vbNullString

    Set DesignerStub = New LinelistSpecsDesignerTranslationStub
    Set MessageStub = New LinelistSpecsTranslationStub
    MessageStub.Initialise
    DesignerStub.UseTranslationObject MessageStub

    Specs.TestAssignMain MainStub
    Specs.TestAssignDesignerTranslations DesignerStub
End Sub

'@section Tests
'===============================================================================

'@TestMethod("LinelistSpecs")
Private Sub TestCreateFailsWhenWorksheetMissing()
    Dim tempBook As Workbook
    Dim errNumber As Long

    Set tempBook = TestHelpers.NewWorkbook()
    PrepareSpecificationWorkbook tempBook, LinelistSpecConstants.SHEET_CHOICES

    On Error Resume Next
        LinelistSpecs.Create tempBook
        errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual ProjectError.ElementNotFound, errNumber, _
                     "Create should fail fast when a required worksheet is missing"

    TestHelpers.DeleteWorkbook tempBook
End Sub

'@TestMethod("LinelistSpecs")
Private Sub TestDictionaryIsCached()
    Dim dictOnce As ILLdictionary
    Dim dictTwice As ILLdictionary

    On Error GoTo Fail
    Set dictOnce = Specs.Dictionary
    Set dictTwice = Specs.Dictionary

    Assert.IsTrue (dictOnce Is dictTwice), _
                  "Dictionary lookup should reuse the cached instance"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestDictionaryIsCached"
End Sub

'@TestMethod("LinelistSpecs")
Private Sub TestResetCachesInvalidatesDictionary()
    Dim firstDict As ILLdictionary
    Dim newDict As ILLdictionary

    On Error GoTo Fail
    Set firstDict = Specs.Dictionary
    Specs.ResetCaches
    Set newDict = Specs.Dictionary

    Assert.IsFalse (firstDict Is newDict), _
                  "ResetCaches must drop the dictionary cache so a new instance is produced"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestResetCachesInvalidatesDictionary"
End Sub

'@TestMethod("LinelistSpecs")
Private Sub TestPrepareRequiresSetupPath()
    Dim errNumber As Long

    On Error Resume Next
        Specs.Prepare
        errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual ProjectError.ElementNotFound, errNumber, _
                     "Prepare should fail fast when the setup path is missing"
End Sub

'@TestMethod("LinelistSpecs")
Private Sub TestExportRequiresDestinationWorkbook()
    Dim errNumber As Long

    On Error Resume Next
        Specs.Export Nothing
        errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual ProjectError.ObjectNotInitialized, errNumber, _
                     "Export should enforce providing a destination workbook"
End Sub

'@section Helpers
'===============================================================================
Private Sub PrepareSpecificationWorkbook(ByVal targetBook As Workbook, _
                                         Optional ByVal excludeSheet As String = vbNullString)

    Dim requiredSheets As Variant
    Dim idx As Long
    Dim sheetName As String
    Dim hostSheet As Worksheet

    requiredSheets = Array( _
        LinelistSpecConstants.SHEET_DICTIONARY, _
        LinelistSpecConstants.SHEET_CHOICES, _
        LinelistSpecConstants.SHEET_GEO, _
        LinelistSpecConstants.SHEET_PASSWORDS, _
        LinelistSpecConstants.SHEET_FORMULAS, _
        LinelistSpecConstants.SHEET_TRANSLATIONS_DICT, _
        LinelistSpecConstants.SHEET_TRANSLATIONS_LL, _
        LinelistSpecConstants.SHEET_ANALYSIS, _
        LinelistSpecConstants.SHEET_EXPORT, _
        LinelistSpecConstants.SHEET_FORMAT, _
        LinelistSpecConstants.SHEET_MAIN, _
        LinelistSpecConstants.SHEET_DESIGNER_TRANSLATION)

    For idx = LBound(requiredSheets) To UBound(requiredSheets)
        sheetName = CStr(requiredSheets(idx))
        If StrComp(sheetName, excludeSheet, vbTextCompare) = 0 Then
            On Error Resume Next
                targetBook.Worksheets(sheetName).Delete
            On Error GoTo 0
        Else
            Set hostSheet = TestHelpers.EnsureWorksheet(sheetName, targetBook)
            hostSheet.Cells.Clear
        End If
    Next idx

    If StrComp(LinelistSpecConstants.SHEET_FORMAT, excludeSheet, vbTextCompare) <> 0 Then
        SeedFormatSheet targetBook.Worksheets(LinelistSpecConstants.SHEET_FORMAT)
    End If

    If StrComp(LinelistSpecConstants.SHEET_DICTIONARY, excludeSheet, vbTextCompare) <> 0 Then
        SeedDictionarySheet targetBook.Worksheets(LinelistSpecConstants.SHEET_DICTIONARY)
    End If

    If StrComp(LinelistSpecConstants.SHEET_CHOICES, excludeSheet, vbTextCompare) <> 0 Then
        SeedChoicesSheet targetBook.Worksheets(LinelistSpecConstants.SHEET_CHOICES)
    End If

    If StrComp(LinelistSpecConstants.SHEET_EXPORT, excludeSheet, vbTextCompare) <> 0 Then
        SeedExportsSheet targetBook.Worksheets(LinelistSpecConstants.SHEET_EXPORT)
    End If
End Sub

Private Sub SeedFormatSheet(ByVal formatSheet As Worksheet)
    formatSheet.Cells.Clear
    formatSheet.Range("A1").Value = TEST_DESIGN_NAME

    On Error Resume Next
        formatSheet.Names(LinelistSpecConstants.RANGE_DESIGN_TYPE).Delete
    On Error GoTo 0

    formatSheet.Names.Add Name:=LinelistSpecConstants.RANGE_DESIGN_TYPE, _
                          RefersTo:=formatSheet.Range("A1")
End Sub

Private Sub SeedDictionarySheet(ByVal dictSheet As Worksheet)
    dictSheet.Cells.Clear
    TestHelpers.WriteRow dictSheet.Range("A1"), "variable name", "control", "control details"
    TestHelpers.WriteRow dictSheet.Range("A2"), "var_choice", "choice_manual", "list_manual"
End Sub

Private Sub SeedChoicesSheet(ByVal choicesSheet As Worksheet)
    choicesSheet.Cells.Clear
    TestHelpers.WriteRow choicesSheet.Range("A1"), "list name", "name", "label"
    TestHelpers.WriteRow choicesSheet.Range("A2"), "list_manual", "option_a", "Option A"
End Sub

Private Sub SeedExportsSheet(ByVal exportsSheet As Worksheet)
    exportsSheet.Cells.Clear
    TestHelpers.WriteRow exportsSheet.Range("A1"), "export name"
    TestHelpers.WriteRow exportsSheet.Range("A2"), "default"
End Sub

