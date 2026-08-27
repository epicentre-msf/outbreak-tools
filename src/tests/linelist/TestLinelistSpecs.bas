Attribute VB_Name = "TestLinelistSpecs"
Attribute VB_Description = "Tests for the LinelistSpecs facade"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the LinelistSpecs facade")

'@description
'Drives LinelistSpecs, the facade the whole generation path runs through. This
'module sat at the src/tests/ root with no folder to be registered under, so
'its tests had never executed and the class had never been compiled by the
'harness.
'
'The fixture is a designer workbook carrying the seven sheets
'RequiredSheetNames() asks for. The four pass-through sheets are absent from it
'on purpose, which is what a designer looks like after the setup-to-linelist
'migration: InitTransfer imports Dictionary, Choices, Analysis and Exports
'straight from the setup file into the linelist.
'@depends LinelistSpecs, CustomTest, LLdictionary

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

Private Const TEST_DESIGN_NAME As String = "UnitTestDesign"

Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_PASSWORDS As String = "__pass"
Private Const SHEET_FORMULAS As String = "__formula"
Private Const SHEET_TRANSLATIONS_LL As String = "LinelistTranslation"
Private Const SHEET_FORMAT As String = "__formatter"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DESIGNER_TRANSLATION As String = "DesignerTranslation"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const RANGE_DESIGN_TYPE As String = "DESIGNTYPE"

Private Assert As CustomTest
Private SpecsWorkbook As Workbook
Private Specs As LinelistSpecs

'@section Module lifecycle
'===============================================================================

'@sub-title Build the designer fixture and the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestLinelistSpecs"

    Set SpecsWorkbook = NewWorkbook()
    PrepareSpecificationWorkbook SpecsWorkbook
    Set Specs = LinelistSpecs.Create(SpecsWorkbook)
End Sub

'@sub-title Print results and drop the fixture workbook.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    Set Specs = Nothing
    DeleteWorkbook SpecsWorkbook
    Set SpecsWorkbook = Nothing
    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Drop every cached collaborator before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    If Not Specs Is Nothing Then Specs.ResetCaches
End Sub

'@sub-title Flush assert state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Creation
'===============================================================================

'@sub-title A designer missing one required sheet is refused at creation.
'@TestMethod("LinelistSpecs")
Public Sub TestCreateFailsWhenWorksheetMissing()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestCreateFailsWhenWorksheetMissing"
    On Error GoTo TestFail

    Dim tempBook As Workbook
    Dim errNumber As Long

    Set tempBook = NewWorkbook()
    PrepareSpecificationWorkbook tempBook, SHEET_GEO

    On Error Resume Next
        LinelistSpecs.Create tempBook
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ElementNotFound), errNumber, _
                    "Create refuses a designer workbook with no Geo worksheet"

    DeleteWorkbook tempBook
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateFailsWhenWorksheetMissing", _
                         Err.Number, Err.Description
End Sub

'@sub-title A Nothing workbook is refused at creation.
'@TestMethod("LinelistSpecs")
Public Sub TestCreateRejectsANothingWorkbook()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestCreateRejectsANothingWorkbook"
    On Error GoTo TestFail

    Dim errNumber As Long

    On Error Resume Next
        LinelistSpecs.Create Nothing
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "Create refuses a Nothing workbook and names the reason"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsANothingWorkbook", Err.Number, Err.Description
End Sub

'@sub-title A facade built with no Application runs in the current one.
'@TestMethod("LinelistSpecs")
Public Sub TestCreateDefaultsToTheRunningApplication()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestCreateDefaultsToTheRunningApplication"
    On Error GoTo TestFail

    Dim facade As LinelistSpecs

    Set facade = LinelistSpecs.Create(SpecsWorkbook)

    Assert.IsTrue facade.HostApplication Is Application, _
                  "The host of a facade built with no Application is the running one"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateDefaultsToTheRunningApplication", _
                         Err.Number, Err.Description
End Sub

'@sub-title The Application handed to Create is the one the facade keeps.
'@TestMethod("LinelistSpecs")
Public Sub TestCreateKeepsTheApplicationItWasGiven()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestCreateKeepsTheApplicationItWasGiven"
    On Error GoTo TestFail

    Dim facade As LinelistSpecs
    Dim hostApp As Application

    Set hostApp = SpecsWorkbook.Application
    Set facade = LinelistSpecs.Create(SpecsWorkbook, hostApp)

    Assert.IsTrue facade.HostApplication Is hostApp, _
                  "HostApplication answers the Application Create was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateKeepsTheApplicationItWasGiven", _
                         Err.Number, Err.Description
End Sub

'@section Caching
'===============================================================================

'@sub-title The dictionary handed in is given back on every read.
'@TestMethod("LinelistSpecs")
Public Sub TestDictionaryIsCached()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestDictionaryIsCached"
    On Error GoTo TestFail

    Dim seeded As LLdictionary
    Dim dictOnce As LLdictionary
    Dim dictTwice As LLdictionary

    Set seeded = BuildDictionary()
    Specs.TestAssignDictionary seeded

    Set dictOnce = Specs.Dictionary
    Set dictTwice = Specs.Dictionary

    Assert.IsTrue (dictOnce Is seeded), "Dictionary gives back the instance it was handed"
    Assert.IsTrue (dictOnce Is dictTwice), "Dictionary gives the same instance on every read"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDictionaryIsCached", Err.Number, Err.Description
End Sub

'@sub-title ResetCaches drops the dictionary the instance was holding.
'@details
'After the reset the facade holds no dictionary, and reading one before Prepare
'has imported the setup file raises. That guard is what replaced the designer
'side loader.
'@TestMethod("LinelistSpecs")
Public Sub TestResetCachesInvalidatesDictionary()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestResetCachesInvalidatesDictionary"
    On Error GoTo TestFail

    Dim seeded As LLdictionary
    Dim afterReset As LLdictionary
    Dim errNumber As Long

    Set seeded = BuildDictionary()
    Specs.TestAssignDictionary seeded

    Assert.IsTrue (Specs.Dictionary Is seeded), "The seeded dictionary is in place"

    Specs.ResetCaches

    On Error Resume Next
        Set afterReset = Specs.Dictionary
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), errNumber, _
                    "Reading a dictionary after ResetCaches names Prepare as the missing step"
    Assert.IsTrue (afterReset Is Nothing), "ResetCaches drops the dictionary the instance held"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestResetCachesInvalidatesDictionary", _
                         Err.Number, Err.Description
End Sub

'@section Prepare
'===============================================================================

'@sub-title Prepare refuses to run without a setup file path.
'@details
'The check has to come before the temporary folder and the output workbook are
'created, or an empty path costs a folder on disk and an orphaned workbook
'before InitTransfer gets to say the same thing.
'@TestMethod("LinelistSpecs")
Public Sub TestPrepareRequiresSetupPath()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestPrepareRequiresSetupPath"
    On Error GoTo TestFail

    Dim errNumber As Long
    Dim openBooks As Long

    openBooks = Workbooks.Count

    On Error Resume Next
        Specs.Prepare vbNullString
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "Prepare refuses to run with no setup file path"
    Assert.AreEqual CLng(openBooks), CLng(Workbooks.Count), _
                    "The refusal leaves no output workbook open"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPrepareRequiresSetupPath", _
                         Err.Number, Err.Description
End Sub

'@sub-title Prepared answers False until Prepare has run.
'@details
'Linelist reads this property before it builds anything. An interface fold
'deleted it and the pair stopped compiling, which is how the generation path
'came to be broken for eleven days.
'@TestMethod("LinelistSpecs")
Public Sub TestPreparedIsFalseBeforePrepare()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestPreparedIsFalseBeforePrepare"
    On Error GoTo TestFail

    Assert.IsFalse Specs.Prepared(), "A fresh facade has not been prepared"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPreparedIsFalseBeforePrepare", _
                         Err.Number, Err.Description
End Sub

'@section The pass-through collaborators
'===============================================================================

'@sub-title Every pass-through collaborator raises before Prepare has run.
'@details
'These four live on the linelist workbook and InitTransfer is what sets
'them. Reading one on a fresh facade used to resolve a designer worksheet that
'the designer no longer carries, so the caller got either stale data or a
'message pointing at the wrong workbook.
'@TestMethod("LinelistSpecs")
Public Sub TestPassThroughAccessorsNeedPrepare()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestPassThroughAccessorsNeedPrepare"
    On Error GoTo TestFail

    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), ErrorOfDictionary(), _
                    "Dictionary names Prepare as the missing step"
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), ErrorOfChoices(), _
                    "Choices names Prepare as the missing step"
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), ErrorOfExports(), _
                    "ExportObject names Prepare as the missing step"
    Assert.AreEqual CLng(ProjectError.ErrorUnexpectedState), ErrorOfAnalysis(), _
                    "AnalysisObject names Prepare as the missing step"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestPassThroughAccessorsNeedPrepare", _
                         Err.Number, Err.Description
End Sub

'@section Configuration values
'===============================================================================

'@sub-title An unknown configuration tag raises.
'@details
'The debug password tag pointed at RNG_RNG_LLPassword, a name that existed
'nowhere, and every generated linelist got an empty debug password with no
'signal. A silent empty answer is what hid it.
'@TestMethod("LinelistSpecs")
Public Sub TestValueRejectsUnknownTag()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestValueRejectsUnknownTag"
    On Error GoTo TestFail

    Dim errNumber As Long
    Dim answer As String

    On Error Resume Next
        answer = Specs.Value("nosuchtag")
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A tag the class does not know raises and names itself"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueRejectsUnknownTag", Err.Number, Err.Description
End Sub

'@sub-title An empty tag answers an empty string.
'@TestMethod("LinelistSpecs")
Public Sub TestValueOfEmptyTagIsEmpty()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestValueOfEmptyTagIsEmpty"
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, Specs.Value(vbNullString), _
                    "An empty tag asks for nothing and answers nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestValueOfEmptyTagIsEmpty", Err.Number, Err.Description
End Sub

'@sub-title A known tag whose named range is missing answers an empty string.
'@details
'The Main worksheet of the fixture carries no named range at all, so every
'range-backed tag lands here. This is the behaviour a designer relies on while
'the user is still filling the form in.
'@TestMethod("LinelistSpecs")
Public Sub TestKnownTagWithNoRangeIsEmpty()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestKnownTagWithNoRangeIsEmpty"
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, Specs.Value("lldir"), _
                    "A known tag whose range is absent answers an empty string"
    Assert.AreEqual vbNullString, Specs.Value("llname"), _
                    "The linelist name reads empty on a form nobody has filled in"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestKnownTagWithNoRangeIsEmpty", _
                         Err.Number, Err.Description
End Sub

'@sub-title The epidemiological week start falls back to week one.
'@TestMethod("LinelistSpecs")
Public Sub TestEpiWeekStartHasADefault()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestEpiWeekStartHasADefault"
    On Error GoTo TestFail

    Assert.AreEqual "1", Specs.Value("epiweekstart"), _
                    "The week start defaults to 1 when the designer holds no value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEpiWeekStartHasADefault", Err.Number, Err.Description
End Sub

'@sub-title The debug password reads the range the designer defines.
'@details
'This is the one tag with a measured answer. The designer workbook defines
'RNG_LLPassword on its Main worksheet, and the tag table asked for
'RNG_RNG_LLPassword, so the value never arrived.
'@TestMethod("LinelistSpecs")
Public Sub TestDebugPasswordReadsTheMainSheetRange()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestDebugPasswordReadsTheMainSheetRange"
    On Error GoTo TestFail

    Dim mainSheet As Worksheet

    Set mainSheet = SpecsWorkbook.Worksheets(SHEET_MAIN)
    mainSheet.Range("B2").Value = "secret-debug"

    On Error Resume Next
        mainSheet.Names("RNG_LLPassword").Delete
    On Error GoTo 0
    mainSheet.Names.Add Name:="RNG_LLPassword", RefersTo:=mainSheet.Range("B2")

    On Error GoTo TestFail
    Assert.AreEqual "secret-debug", Specs.Value("debugpassword"), _
                    "The debug password comes from RNG_LLPassword on the Main worksheet"

    mainSheet.Names("RNG_LLPassword").Delete
    mainSheet.Range("B2").ClearContents
    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestDebugPasswordReadsTheMainSheetRange", _
                         Err.Number, Err.Description
End Sub

'@section Temporary sheet names
'===============================================================================

'@sub-title Every temporary sheet scope answers its worksheet name.
'@details
'Metadata is in this list on purpose. LLGeo creates a real Metadata worksheet
'in the linelist, and the preserved-name list is what stops the dictionary
'generating a data-entry sheet that collides with it.
'@TestMethod("LinelistSpecs")
Public Sub TestTemporarySheetNames()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestTemporarySheetNames"
    On Error GoTo TestFail

    Assert.AreEqual "__temp", Specs.TemporarySheetName(TempSheetSingle), _
                    "The single temporary sheet is __temp"
    Assert.AreEqual "__dropdown_lists", Specs.TemporarySheetName(TempSheetList), _
                    "The list sheet is __dropdown_lists"
    Assert.AreEqual "Metadata", Specs.TemporarySheetName(TempSheetMetadata), _
                    "The geo metadata sheet is Metadata"
    Assert.AreEqual "__ana_tabnames", Specs.TemporarySheetName(TempSheetAnalysis), _
                    "The analysis registry sheet is __ana_tabnames"
    Assert.AreEqual "__import_rep", Specs.TemporarySheetName(TempSheetImport), _
                    "The import report sheet is __import_rep"
    Assert.AreEqual "__spatial_tables", Specs.TemporarySheetName(TempSheetSpatial), _
                    "The spatial sheet is __spatial_tables"
    Assert.AreEqual "__show_hide", Specs.TemporarySheetName(TempSheetShowHide), _
                    "The show and hide sheet is __show_hide"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTemporarySheetNames", Err.Number, Err.Description
End Sub

'@sub-title An unknown temporary sheet scope raises.
'@TestMethod("LinelistSpecs")
Public Sub TestTemporarySheetNameRejectsUnknownScope()
    CustomTestSetTitles Assert, "LinelistSpecs", "TestTemporarySheetNameRejectsUnknownScope"
    On Error GoTo TestFail

    Dim errNumber As Long
    Dim answer As String

    On Error Resume Next
        answer = Specs.TemporarySheetName(200)
        errNumber = Err.Number
    On Error GoTo 0

    On Error GoTo TestFail
    Assert.AreEqual CLng(ProjectError.InvalidArgument), errNumber, _
                    "A scope outside the enum raises and names itself"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTemporarySheetNameRejectsUnknownScope", _
                         Err.Number, Err.Description
End Sub

'@section Fixture helpers
'===============================================================================

'@sub-title The error number a Dictionary read gives on a fresh facade.
'@return Long. The error number, or zero when the read succeeded.
Private Function ErrorOfDictionary() As Long
    Dim answer As LLdictionary
    On Error Resume Next
        Set answer = Specs.Dictionary
        ErrorOfDictionary = Err.Number
    On Error GoTo 0
End Function

'@sub-title The error number a Choices read gives on a fresh facade.
'@return Long. The error number, or zero when the read succeeded.
Private Function ErrorOfChoices() As Long
    Dim answer As LLChoices
    On Error Resume Next
        Set answer = Specs.Choices
        ErrorOfChoices = Err.Number
    On Error GoTo 0
End Function

'@sub-title The error number an ExportObject read gives on a fresh facade.
'@return Long. The error number, or zero when the read succeeded.
Private Function ErrorOfExports() As Long
    Dim answer As LLExport
    On Error Resume Next
        Set answer = Specs.ExportObject
        ErrorOfExports = Err.Number
    On Error GoTo 0
End Function

'@sub-title The error number an AnalysisObject read gives on a fresh facade.
'@return Long. The error number, or zero when the read succeeded.
Private Function ErrorOfAnalysis() As Long
    Dim answer As Analysis
    On Error Resume Next
        Set answer = Specs.AnalysisObject
        ErrorOfAnalysis = Err.Number
    On Error GoTo 0
End Function

'@sub-title Build a dictionary on a worksheet of the fixture workbook.
'@return LLdictionary. A dictionary over a two-row variables table.
Private Function BuildDictionary() As LLdictionary
    Dim dictSheet As Worksheet

    Set dictSheet = EnsureWorksheet(SHEET_DICTIONARY, SpecsWorkbook)
    WriteRow dictSheet.Range("A1"), "variable name", "control", "control details"
    WriteRow dictSheet.Range("A2"), "var_choice", "choice_manual", "list_manual"

    Set BuildDictionary = LLdictionary.Create(dictSheet, 1, 1, 1)
End Function

'@sub-title Seed a workbook with the worksheets a designer has to carry.
'@details
'Only the seven sheets RequiredSheetNames() asks for are created. Dictionary,
'Choices, Analysis and Exports are left out because InitTransfer imports
'them straight from the setup file into the linelist.
'@param targetBook Workbook. The workbook to seed.
'@param excludeSheet String. One sheet name to leave out, for the refusal test.
Private Sub PrepareSpecificationWorkbook(ByVal targetBook As Workbook, _
                                         Optional ByVal excludeSheet As String = vbNullString)

    Dim requiredSheets As Variant
    Dim idx As Long
    Dim sheetName As String
    Dim hostSheet As Worksheet

    requiredSheets = Array( _
        SHEET_GEO, _
        SHEET_PASSWORDS, _
        SHEET_FORMULAS, _
        SHEET_TRANSLATIONS_LL, _
        SHEET_FORMAT, _
        SHEET_MAIN, _
        SHEET_DESIGNER_TRANSLATION)

    For idx = LBound(requiredSheets) To UBound(requiredSheets)
        sheetName = CStr(requiredSheets(idx))
        If StrComp(sheetName, excludeSheet, vbTextCompare) = 0 Then
            On Error Resume Next
                targetBook.Worksheets(sheetName).Delete
            On Error GoTo 0
        Else
            Set hostSheet = EnsureWorksheet(sheetName, targetBook)
            hostSheet.Cells.Clear
        End If
    Next idx

    If StrComp(SHEET_FORMAT, excludeSheet, vbTextCompare) <> 0 Then
        SeedFormatSheet targetBook.Worksheets(SHEET_FORMAT)
    End If

    If StrComp(SHEET_DESIGNER_TRANSLATION, excludeSheet, vbTextCompare) <> 0 Then
        SeedDesignerTranslationSheet targetBook.Worksheets(SHEET_DESIGNER_TRANSLATION)
    End If
End Sub

'@sub-title Seed the four ListObjects DesignerTranslation.Create requires (ENG)
'@param designerSheet Worksheet. The DesignerTranslation worksheet.
Private Sub SeedDesignerTranslationSheet(ByVal designerSheet As Worksheet)
    designerSheet.Cells.Clear
    AddTradTable designerSheet, designerSheet.Range("A1"), "T_tradMsg", _
        Array(Array("tag", "ENG"), Array("MSG_Info", "Information"))
    AddTradTable designerSheet, designerSheet.Range("D1"), "T_tradShape", _
        Array(Array("tag", "ENG"), Array("shp_title", "Designer"))
    AddTradTable designerSheet, designerSheet.Range("G1"), "T_tradRange", _
        Array(Array("tag", "ENG"), Array("RNG_DesignerTitle", "Designer Title"))
    AddTradTable designerSheet, designerSheet.Range("J1"), "T_tradDrop", _
        Array(Array("tag", "ENG"), Array("drp_choice", "list_values"))
End Sub

'@sub-title Write one translation table and name it.
'@param sh Worksheet. The host worksheet.
'@param startCell Range. The top left cell of the table.
'@param tableName String. The name to give the ListObject.
'@param tableRows Variant. An array of row arrays, header first.
Private Sub AddTradTable(ByVal sh As Worksheet, ByVal startCell As Range, _
                         ByVal tableName As String, ByVal tableRows As Variant)
    Dim matrix As Variant
    Dim dataRange As Range
    Dim lo As ListObject

    matrix = RowsToMatrix(tableRows)
    WriteMatrix startCell, matrix
    Set dataRange = startCell.Resize(UBound(matrix, 1), UBound(matrix, 2))
    Set lo = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = tableName
End Sub

'@sub-title Seed the format worksheet with a DESIGNTYPE named range.
'@param formatSheet Worksheet. The __formatter worksheet.
Private Sub SeedFormatSheet(ByVal formatSheet As Worksheet)
    formatSheet.Cells.Clear
    formatSheet.Range("A1").Value = TEST_DESIGN_NAME

    On Error Resume Next
        formatSheet.Names(RANGE_DESIGN_TYPE).Delete
    On Error GoTo 0

    formatSheet.Names.Add Name:=RANGE_DESIGN_TYPE, _
                          RefersTo:=formatSheet.Range("A1")
End Sub
