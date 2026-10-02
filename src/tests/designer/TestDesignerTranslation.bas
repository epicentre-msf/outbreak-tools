Attribute VB_Name = "TestDesignerTranslation"
Attribute VB_Description = "Unit tests for DesignerTranslation class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates DesignerTranslation: factory requirements, the persisted language code, the cached translation objects, and the shape, range and dropdown application on a designer worksheet.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private TradSheet As Worksheet
Private MainSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"

'The hidden name DesignerTranslation persists the language code under
Private Const LANG_HIDDEN_NAME As String = "TAG_DES_LANG"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDesignerTranslation"
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
    On Error GoTo 0
    Set Assert = Nothing
    RestoreApp
End Sub


'@section Test lifecycle
'===============================================================================
'@TestInitialize
Public Sub TestInitialize()
    BusyApp

    Set FixtureWorkbook = NewWorkbook
    Set TradSheet = EnsureWorksheet("DesignerTradTables", FixtureWorkbook)
    Set MainSheet = EnsureWorksheet("Main", FixtureWorkbook)
    SeedTradTables TradSheet
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set MainSheet = Nothing
    Set TradSheet = Nothing
    Set FixtureWorkbook = Nothing

    RestoreApp
End Sub


'@section Factory Tests
'===============================================================================
'@TestMethod("DesignerTranslation.Factory")
Public Sub TestCreateRefusesNothingWorksheet()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestCreateRefusesNothingWorksheet"
    On Error GoTo Fail

    Dim raisedNumber As Long
    Dim subject As DesignerTranslation

    On Error Resume Next
    Set subject = DesignerTranslation.Create(Nothing)
    raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.IsTrue raisedNumber <> 0, "Create should refuse a Nothing worksheet."
    Assert.IsNothing subject, "No instance should come back from a refused Create."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCreateRefusesNothingWorksheet", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Factory")
Public Sub TestCreateRefusesMissingTable()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestCreateRefusesMissingTable"
    On Error GoTo Fail

    'Arrange: a sheet carrying three of the four required tables
    TradSheet.ListObjects("T_tradDrop").Delete

    Dim raisedNumber As Long
    Dim subject As DesignerTranslation

    On Error Resume Next
    Set subject = DesignerTranslation.Create(TradSheet)
    raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.IsTrue raisedNumber <> 0, "Create should refuse a sheet with a missing translation table."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestCreateRefusesMissingTable", Err.Number, Err.Description
End Sub


'@section Language Code Tests
'===============================================================================
'@TestMethod("DesignerTranslation.Language")
Public Sub TestTransObjectAnswersNothingWithoutLanguage()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTransObjectAnswersNothingWithoutLanguage"
    On Error GoTo Fail

    'Arrange: a fresh sheet holds no persisted language code
    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act and assert
    Assert.IsNothing subject.TransObject(), _
                     "TransObject should answer Nothing when no language code is stored."
    Assert.AreEqual vbNullString, subject.TranslatedValue("MSG_Info"), _
                    "TranslatedValue should answer an empty string when no language code is stored."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTransObjectAnswersNothingWithoutLanguage", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Language")
Public Sub TestTranslateDesignerPersistsLanguageCode()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTranslateDesignerPersistsLanguageCode"
    On Error GoTo Fail

    'Arrange
    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act: the language string carries the code before the dash
    subject.TranslateDesigner MainSheet, "FRA - Francais"

    'Assert: the code lands on the worksheet-level hidden name
    Dim store As HiddenNames
    Set store = HiddenNames.Create(TradSheet)
    Assert.IsTrue store.HasName(LANG_HIDDEN_NAME), _
                  "TranslateDesigner should create the language hidden name."
    Assert.AreEqual "FRA", store.ValueAsString(LANG_HIDDEN_NAME), _
                    "The persisted code should be the part before the dash."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTranslateDesignerPersistsLanguageCode", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Language")
Public Sub TestStoredLanguageDrivesAFreshInstance()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestStoredLanguageDrivesAFreshInstance"
    On Error GoTo Fail

    'Arrange: one instance persists the code
    Dim first As DesignerTranslation
    Set first = DesignerTranslation.Create(TradSheet)
    first.TranslateDesigner MainSheet, "FRA - Francais"

    'Act: a fresh instance over the same sheet reads the stored code
    Dim second As DesignerTranslation
    Set second = DesignerTranslation.Create(TradSheet)

    'Assert
    Assert.AreEqual "Informations FR", second.TranslatedValue("MSG_Info"), _
                    "A fresh instance should translate with the stored language code."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestStoredLanguageDrivesAFreshInstance", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Language")
Public Sub TestLanguageChangeRebuildsCaches()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestLanguageChangeRebuildsCaches"
    On Error GoTo Fail

    'Arrange
    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act and assert: the first language answers its own rows
    subject.TranslateDesigner MainSheet, "ENG - English"
    Assert.AreEqual "Information", subject.TranslatedValue("MSG_Info"), _
                    "The first language should answer its own rows."

    'Act and assert: a language change drops the caches and reads the new rows
    subject.TranslateDesigner MainSheet, "FRA - Francais"
    Assert.AreEqual "Informations FR", subject.TranslatedValue("MSG_Info"), _
                    "A language change should rebuild the caches on the new language."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestLanguageChangeRebuildsCaches", Err.Number, Err.Description
End Sub


'@section Translation Object Tests
'===============================================================================
'@TestMethod("DesignerTranslation.TransObject")
Public Sub TestTransObjectAnswersEachScope()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTransObjectAnswersEachScope"
    On Error GoTo Fail

    'Arrange
    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)
    subject.TranslateDesigner MainSheet, "ENG - English"

    'Act and assert: each scope translates from its own table
    Assert.AreEqual "Information", _
                    subject.TransObject(DesignerTranslationOfMessages).TranslatedValue("MSG_Info"), _
                    "The messages scope should read T_tradMsg."
    Assert.AreEqual "Designer", _
                    subject.TransObject(DesignerTranslationOfShapes).TranslatedValue("shp_title"), _
                    "The shapes scope should read T_tradShape."
    Assert.AreEqual "Designer Title", _
                    subject.TransObject(DesignerTranslationOfRanges).TranslatedValue("RNG_DesignerTitle"), _
                    "The ranges scope should read T_tradRange."
    Assert.AreEqual "list_values", _
                    subject.TransObject(DesignerTranslationOfDropdowns).TranslatedValue("drp_choice"), _
                    "The dropdowns scope should read T_tradDrop."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTransObjectAnswersEachScope", Err.Number, Err.Description
End Sub


'@section Translation Application Tests
'===============================================================================
'@TestMethod("DesignerTranslation.Apply")
Public Sub TestTranslateDesignerUpdatesShapes()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTranslateDesignerUpdatesShapes"
    On Error GoTo Fail

    'Arrange: a shape named in T_tradShape, with placeholder text
    Dim shp As Shape
    Set shp = MainSheet.Shapes.AddShape(msoShapeRectangle, 10, 10, 120, 30)
    shp.Name = "shp_title"
    shp.TextFrame.Characters.Text = "placeholder"

    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act
    subject.TranslateDesigner MainSheet, "ENG - English"

    'Assert
    Assert.AreEqual "Designer", shp.TextFrame.Characters.Text, _
                    "The shape text should carry the translated value."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTranslateDesignerUpdatesShapes", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Apply")
Public Sub TestTranslateDesignerWritesRangeValues()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTranslateDesignerWritesRangeValues"
    On Error GoTo Fail

    'Arrange: the named range T_tradRange names, with placeholder content
    FixtureWorkbook.Names.Add Name:="RNG_DesignerTitle", RefersTo:=MainSheet.Range("B1")
    MainSheet.Range("B1").Value = "placeholder"

    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act
    subject.TranslateDesigner MainSheet, "ENG - English"

    'Assert
    Assert.AreEqual "Designer Title", CStr(MainSheet.Range("B1").Value), _
                    "The named range should carry the translated value."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTranslateDesignerWritesRangeValues", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Apply")
Public Sub TestTranslateDesignerAppliesDropdownValidation()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTranslateDesignerAppliesDropdownValidation"
    On Error GoTo Fail

    'Arrange: the dropdown cell and the list its validation points at
    FixtureWorkbook.Names.Add Name:="drp_choice", RefersTo:=MainSheet.Range("C1")
    MainSheet.Range("C1").Value = "stale choice"

    MainSheet.Range("E1").Value = "alpha"
    MainSheet.Range("E2").Value = "beta"
    FixtureWorkbook.Names.Add Name:="list_values", RefersTo:=MainSheet.Range("E1:E2")

    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act
    subject.TranslateDesigner MainSheet, "ENG - English"

    'Assert: a list validation stands on the cell and the content is cleared
    Assert.AreEqual CLng(xlValidateList), CLng(MainSheet.Range("C1").Validation.Type), _
                    "The dropdown cell should carry a list validation."
    Assert.IsTrue InStr(1, MainSheet.Range("C1").Validation.Formula1, "list_values") > 0, _
                  "The validation formula should point at the translated list name."
    Assert.AreEqual vbNullString, CStr(MainSheet.Range("C1").Value), _
                    "The dropdown cell content should be cleared after the validation lands."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTranslateDesignerAppliesDropdownValidation", Err.Number, Err.Description
End Sub

'@TestMethod("DesignerTranslation.Apply")
Public Sub TestTranslateDesignerWithoutLanguageLeavesSheetAlone()
    CustomTestSetTitles Assert, "DesignerTranslation", "TestTranslateDesignerWithoutLanguageLeavesSheetAlone"
    On Error GoTo Fail

    'Arrange: no stored code and an empty language string
    FixtureWorkbook.Names.Add Name:="RNG_DesignerTitle", RefersTo:=MainSheet.Range("B1")
    MainSheet.Range("B1").Value = "untouched"

    Dim subject As DesignerTranslation
    Set subject = DesignerTranslation.Create(TradSheet)

    'Act
    subject.TranslateDesigner MainSheet, vbNullString

    'Assert
    Assert.AreEqual "untouched", CStr(MainSheet.Range("B1").Value), _
                    "An empty language with no stored code should leave the sheet alone."

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestTranslateDesignerWithoutLanguageLeavesSheetAlone", Err.Number, Err.Description
End Sub


'@section Test helpers
'===============================================================================

'@sub-title Seed the four designer translation ListObjects with ENG and FRA columns
Private Sub SeedTradTables(ByVal sh As Worksheet)
    sh.Cells.Clear
    AddTradTable sh, sh.Range("A1"), "T_tradMsg", Array( _
        Array("tag", "ENG", "FRA"), _
        Array("MSG_Info", "Information", "Informations FR"), _
        Array("MSG_ChemFich", "File path loaded", "Chemin charge"))
    AddTradTable sh, sh.Range("E1"), "T_tradShape", Array( _
        Array("tag", "ENG", "FRA"), _
        Array("shp_title", "Designer", "Concepteur"))
    AddTradTable sh, sh.Range("I1"), "T_tradRange", Array( _
        Array("tag", "ENG", "FRA"), _
        Array("RNG_DesignerTitle", "Designer Title", "Titre du concepteur"))
    AddTradTable sh, sh.Range("M1"), "T_tradDrop", Array( _
        Array("tag", "ENG", "FRA"), _
        Array("drp_choice", "list_values", "list_values"))
End Sub

'@sub-title Write one translation ListObject from inline rows
Private Sub AddTradTable(ByVal sh As Worksheet, ByVal startCell As Range, _
                         ByVal tableName As String, ByVal rows As Variant)
    Dim matrix As Variant
    Dim dataRange As Range
    Dim lo As ListObject

    matrix = RowsToMatrix(rows)
    WriteMatrix startCell, matrix
    Set dataRange = startCell.Resize(UBound(matrix, 1), UBound(matrix, 2))
    Set lo = sh.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
    lo.Name = tableName
End Sub
