Attribute VB_Name = "TestDiseaseTranslationCache"
Attribute VB_Description = "Tests ensuring DiseaseTranslationCache caches translated headers and respects fallbacks"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests ensuring DiseaseTranslationCache caches translated headers and respects fallbacks")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TRANSLATION_SHEET As String = "CacheTranslationFixture"

Private Assert As ICustomTest
Private Cache As IDiseaseTranslationCache
Private Translations As ITranslationObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseTranslationCache"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheet TRANSLATION_SHEET
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
    Set Cache = Nothing
    Set Translations = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    DeleteWorksheet TRANSLATION_SHEET
    PrepareTranslations
    Set Cache = DiseaseTranslationCache.Create(Translations)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    DeleteWorksheet TRANSLATION_SHEET
    Set Cache = Nothing
    Set Translations = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("DiseaseTranslationCache")
Public Sub TestResolveCachesValues()
    CustomTestSetTitles Assert, "DiseaseTranslationCache", "TestResolveCachesValues"

    Dim resolved As String
    Dim cached As String
    Dim lo As ListObject

    On Error GoTo Fail

    resolved = Cache.Resolve("list name", "List")

    Set lo = Translations.Table
    lo.DataBodyRange.Cells(1, 2).Value = "Modified"

    cached = Cache.Resolve("list name", "List")

    Assert.AreEqual "List Name", resolved, "Initial translation should be returned"
    Assert.AreEqual "List Name", cached, "Cached translation should ignore subsequent changes"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResolveCachesValues", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseTranslationCache")
Public Sub TestResolveFallsBackWhenMissing()
    CustomTestSetTitles Assert, "DiseaseTranslationCache", "TestResolveFallsBackWhenMissing"

    Dim fallback As String

    On Error GoTo Fail

    fallback = Cache.Resolve("unknown_key", "Fallback")
    Assert.AreEqual "Fallback", fallback, "Missing translations should return the supplied fallback"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestResolveFallsBackWhenMissing", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseTranslationCache")
Public Sub TestResolveRaisesOnEmptyKey()
    CustomTestSetTitles Assert, "DiseaseTranslationCache", "TestResolveRaisesOnEmptyKey"

    Dim raisedError As Boolean

    On Error Resume Next
        Cache.Resolve vbNullString, "Fallback"
        raisedError = (Err.Number = ProjectError.InvalidArgument)
        Err.Clear
    On Error GoTo 0

    Assert.IsTrue raisedError, "Resolve should raise when key is empty"
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareTranslations()
    Dim sheet As Worksheet
    Dim header As Variant
    Dim dataRows As Variant
    Dim listRange As Range

    Set sheet = EnsureWorksheet(TRANSLATION_SHEET)
    ClearWorksheet sheet

    header = RowsToMatrix(Array(Array("tag", "ENG")))
    dataRows = RowsToMatrix(Array( _
        Array("list name", "List Name"), _
        Array("label", "Label") _
    ))

    WriteMatrix sheet.Range("A1"), header
    WriteMatrix sheet.Range("A2"), dataRows

    Set listRange = sheet.Range("A1").Resize(UBound(dataRows, 1) + 1, UBound(dataRows, 2))
    sheet.ListObjects.Add SourceType:=xlSrcRange, Source:=listRange, XlListObjectHasHeaders:=xlYes

    Set Translations = TranslationObject.Create(sheet.ListObjects(1), "ENG")
End Sub
