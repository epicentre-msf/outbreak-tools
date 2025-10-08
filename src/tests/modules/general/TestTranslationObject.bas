Attribute VB_Name = "TestTranslationObject"

Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")

Private Const TRANSLATIONNAME As String = "TranslationFixture"
Private Const TRANSLATIONTABLE As String = "tblTranslation"
Private Const TARGETLANGUAGE As String = "French"
Private Const DEFAULTOUTPUTSHEET As String = "testsOutputs"

Private Assert As ICustomTest
Private Translator As ITranslationObject
Private wb As Workbook

'@section Helpers
'===============================================================================

Private Function TranslationHeaders() As Variant
    TranslationHeaders = Array("tag", "English", TARGETLANGUAGE)
End Function

Private Function TranslationRows() As Variant
    TranslationRows = Array( _
        Array("greeting", "Hello", "Bonjour"), _
        Array("farewell", "Goodbye", "Au revoir"), _
        Array("status_ok", "Status is ok", "Le statut est correct"))
End Function

Private Function PrepareTranslationTable() As ListObject
    Dim translationSheet As Worksheet
    Dim headerMatrix As Variant
    Dim dataMatrix As Variant
    Dim translationList As ListObject

    Set translationSheet = EnsureWorksheet(TRANSLATIONNAME)
    ClearWorksheet translationSheet

    headerMatrix = RowsToMatrix(Array(TranslationHeaders()))
    WriteMatrix translationSheet.Cells(1, 1), headerMatrix

    dataMatrix = RowsToMatrix(TranslationRows())
    WriteMatrix translationSheet.Cells(2, 1), dataMatrix

    Set translationList = translationSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                                           Source:=translationSheet.Range("A1").CurrentRegion, _
                                                           XlListObjectHasHeaders:=xlYes)
    translationList.Name = TRANSLATIONTABLE

    Set PrepareTranslationTable = translationList
End Function

Private Sub ResetTranslator()
    Dim translationTable As ListObject
    Set translationTable = PrepareTranslationTable()
    Set Translator = TranslationObject.Create(translationTable, TARGETLANGUAGE)
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set wb = ThisWorkbook
    EnsureWorksheet DEFAULTOUTPUTSHEET
    Set Assert = CustomTest.Create(wb, DEFAULTOUTPUTSHEET)
    ResetTranslator
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    DeleteWorksheet TRANSLATIONNAME
    Assert.PrintResults
    Set Translator = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    BusyApp
    ResetTranslator
End Sub

'@TestCleanup
Public Sub TestCleanup()
    Set Translator = Nothing
    Assert.Flush
End Sub

'@section Tests
'===============================================================================

'@TestMethod("TranslationObject")
Public Sub TestCreateInitialisesTranslation()
    CustomTestSetTitles Assert, "TranslationObject", "TestCreateInitialisesTranslation"
    On Error GoTo Fail

    Assert.IsTrue Translator.ValueExists("greeting"), "Expected greeting tag to exist"
    Assert.AreEqual "Bonjour", Translator.TranslatedValue("greeting"), "Greeting should translate to French"
    Assert.AreEqual "missing_tag", Translator.TranslatedValue("missing_tag"), "Missing tags should return original value"
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateInitialisesTranslation", Err.Number, Err.Description
End Sub

'@TestMethod("TranslationObject")
Public Sub TestTranslateRangeHandlesFormulas()
    On Error GoTo Fail
    CustomTestSetTitles Assert, "TranslationObject", "TestCreateInitialisesTranslation"

    Dim translationSheet As Worksheet

    Set translationSheet = ThisWorkbook.Worksheets(TRANSLATIONNAME)
    translationSheet.Range("E1").Value = "greeting"
    translationSheet.Range("E2").Value = "farewell"
    Translator.TranslateRange translationSheet.Range("E1:E2")
    Assert.AreEqual "Bonjour", translationSheet.Range("E1").Value, "TranslateRange should translate basic values"
    Assert.AreEqual "Au revoir", translationSheet.Range("E2").Value, "TranslateRange should process multiple cells"

    translationSheet.Range("F1").Value = "IF(test,""status_ok"", another_test)"
    Translator.TranslateRange translationSheet.Range("F1"), containsFormula:=True
    Assert.IsTrue InStr(translationSheet.Range("F1").Value, "Le statut est correct") > 0, "Formula segments should be translated"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestTranslateRangeHandlesFormulas"
End Sub
