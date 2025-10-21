Attribute VB_Name = "TestLinelistTranslations"

Option Explicit

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TRANSLATION_SHEET_NAME As String = "LinelistTranslations"
Private Const EXPORT_DICTIONARY_SHEET_NAME As String = "Translation"

'@Folder("CustomTests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As ICustomTest
Private Manager As ILinelistTranslations
Private SourceBook As Workbook
Private TargetBook As Workbook


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestLinelistTranslations"
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
    Set SourceBook = TestHelpers.NewWorkbook
    PrepareSourceWorkbook SourceBook
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not Manager Is Nothing Then
        Set Manager = Nothing
    End If

    If Not TargetBook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook TargetBook
        Set TargetBook = Nothing
    End If

    If Not SourceBook Is Nothing Then
        BusyApp
        TestHelpers.DeleteWorkbook SourceBook
        Set SourceBook = Nothing
    End If
End Sub


'@section Tests
'===============================================================================
'@TestMethod("LinelistTranslations")
Public Sub TestValueRoundTrip()
    CustomTestSetTitles Assert, "LinelistTranslations", "ValueRoundTrip"

    Manager.SetValue "lllanguage", "Français"

    Assert.AreEqual "Français", Manager.Value("lllanguage"), _
                     "SetValue should update the underlying named range"
    Assert.AreEqual "Français", SheetNameValue(SourceBook.Worksheets(TRANSLATION_SHEET_NAME), "RNG_LLLanguage"), _
                     "Name definition should reflect updated value on translation sheet"
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestTransObjectResolvesMessages()
    CustomTestSetTitles Assert, "LinelistTranslations", "TransObjectResolvesMessages"

    Dim translations As ITranslationObject
    Set translations = Manager.TransObject(translationScopeMessages)

    Dim containsFormula As Boolean
    Dim translated As String
    translated = translations.TranslatedValue("HELLO", containsFormula)

    Assert.IsFalse containsFormula, "Message translation should not contain formulas"
    Assert.AreEqual "Hello", translated, "Translated value should match the message table entry"
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestExportToWorkbookCopiesSheets()
    CustomTestSetTitles Assert, "LinelistTranslations", "ExportToWorkbookCopiesSheets"

    Set TargetBook = TestHelpers.NewWorkbook
    Manager.ExportToWorkbook TargetBook, xlSheetVisible

    Dim translationSheet As Worksheet
    Set translationSheet = TargetBook.Worksheets(TRANSLATION_SHEET_NAME)

    Assert.AreEqual xlSheetVisible, translationSheet.Visible, "Translation sheet should respect requested visibility"

    Dim exportTable As ListObject
    Set exportTable = translationSheet.ListObjects(DICTIONARY_TABLE_NAME())

    Assert.AreEqual "Bonjour", exportTable.DataBodyRange.Cells(1, 2).Value, _
                     "Dictionary export should preserve translated values"

    Assert.IsTrue SheetHasName(translationSheet, "LLDictionaryTranslation"), _
                  "Translation sheet should retain dictionary name binding"
    Assert.IsTrue SheetHasName(translationSheet, "RNG_SetupDictionarySheet"), _
                  "Setup dictionary sheet name should be exported with translation sheet"
    Assert.IsFalse WorkbookHasSheet(TargetBook, EXPORT_DICTIONARY_SHEET_NAME), _
                  "Workbook export should not create a separate dictionary sheet"
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestImportSynchronisesDictionary()
    CustomTestSetTitles Assert, "LinelistTranslations", "ImportSynchronisesDictionary"

    Set TargetBook = BuildWorkbookWithInlineDictionary("Hola")

    Manager.ImportFromWorkbook TargetBook

    Dim dictionaryTranslations As ITranslationObject
    Set dictionaryTranslations = Manager.TransObject(translationScopeDictionary)

    Dim containsFormula As Boolean
    Dim translated As String
    translated = dictionaryTranslations.TranslatedValue("HELLO", containsFormula)

    Assert.AreEqual "Hola", translated, "Import should pull dictionary data from translation sheet when present"
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestImportClearsDictionaryWhenMissing()
    CustomTestSetTitles Assert, "LinelistTranslations", "ImportClearsDictionaryWhenMissing"

    Dim blankBook As Workbook
    Set blankBook = BuildWorkbookWithoutDictionary()

    Manager.ImportFromWorkbook blankBook

    Dim dictionaryTable As ListObject
    Set dictionaryTable = SourceBook.Worksheets(TRANSLATION_SHEET_NAME).ListObjects(DICTIONARY_TABLE_NAME())

    Assert.IsTrue DictionaryTableIsEmpty(dictionaryTable), _
                  "Dictionary table should be cleared when import provides no translations"
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestImportFromSetupUpdatesDictionary()
    CustomTestSetTitles Assert, "LinelistTranslations", "ImportFromSetupUpdatesDictionary"

    Dim setupSheet As Worksheet
    Set setupSheet = BuildSetupWorksheet("HELLO", "Hallo")

    Manager.ImportFromSetup setupSheet

    Dim translations As ITranslationObject
    Set translations = Manager.TransObject(translationScopeDictionary)

    Dim containsFormula As Boolean
    Dim translated As String
    translated = translations.TranslatedValue("HELLO", containsFormula)

    Assert.AreEqual "Hallo", translated, "ImportFromSetup should replace dictionary translations"

    BusyApp
    TestHelpers.DeleteWorkbook setupSheet.Parent
    Set setupSheet = Nothing
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestExportSetupRespectsStoredSheetName()
    CustomTestSetTitles Assert, "LinelistTranslations", "ExportSetupRespectsStoredSheetName"

    Dim setupSheet As Worksheet
    Set setupSheet = BuildSetupWorksheet("HELLO", "Ciao")
    Manager.ImportFromSetup setupSheet

    Set TargetBook = TestHelpers.NewWorkbook
    Manager.ExportSetup TargetBook, xlSheetVisible

    Dim exported As Worksheet
    Set exported = TargetBook.Worksheets(setupSheet.Name)

    Assert.AreEqual xlSheetVisible, exported.Visible, "Setup export should respect requested visibility"
    Assert.IsTrue SheetHasName(exported, "LLDictionaryTranslation"), _
                  "Setup export should tag the dictionary listobject with a sheet-level name"
    Assert.IsTrue SheetHasName(exported, "RNG_SetupDictionarySheet"), _
                  "Setup export should preserve the setup sheet named value"

    Dim exportTable As ListObject
    Set exportTable = exported.ListObjects(DICTIONARY_TABLE_NAME())

    Assert.AreEqual "Ciao", exportTable.DataBodyRange.Cells(1, 2).Value, _
                     "Setup export should copy dictionary translations"

    BusyApp
    TestHelpers.DeleteWorkbook setupSheet.Parent
    Set setupSheet = Nothing
End Sub

'@TestMethod("LinelistTranslations")
Public Sub TestCreateRequiresTranslationTables()
    CustomTestSetTitles Assert, "LinelistTranslations", "CreateRequiresTranslationTables"

    Dim faultyBook As Workbook
    Set faultyBook = TestHelpers.NewWorkbook

    Dim translationSheet As Worksheet

    Set translationSheet = faultyBook.Worksheets(1)
    translationSheet.Name = TRANSLATION_SHEET_NAME

    On Error GoTo ExpectFailure
        LinelistTranslations.Create translationSheet
        Assert.Fail "Initialisation should fail when translation tables are missing"
        GoTo CleanUp
ExpectFailure:
    Assert.AreEqual ProjectError.ElementNotFound, Err.Number, _
                     "Missing tables should raise ElementNotFound"
    Err.Clear
CleanUp:
    BusyApp
    TestHelpers.DeleteWorkbook faultyBook
End Sub


'@section Helpers
'===============================================================================
Private Sub PrepareSourceWorkbook(ByVal workbook As Workbook)
    Dim translationSheet As Worksheet

    Set translationSheet = workbook.Worksheets(1)
    translationSheet.Name = TRANSLATION_SHEET_NAME

    SetupTranslationTables translationSheet
    SetupNamedValues translationSheet
    PopulateDictionaryTable translationSheet, "HELLO", "Bonjour"

    Set Manager = LinelistTranslations.Create(translationSheet)
End Sub

Private Sub SetupTranslationTables(ByVal sheet As Worksheet)
    CreateTable sheet, "T_TradLLMsg", "A1:B3", _
                Array("Key", "EN"), Array(Array("HELLO", "Hello"), _
                Array("WORLD", "World"))

    CreateTable sheet, "T_TradLLShapes", "D1:E2", _
                Array("Key", "EN"), Array(Array("CIRCLE", "Circle"))

    CreateTable sheet, "T_TradLLForms", "G1:H2", _
                Array("Key", "EN"), Array(Array("FORM1", "Form One"))

    CreateTable sheet, "T_TradLLRibbon", "J1:K2", _
                Array("Key", "EN"), Array(Array("RIBBON_REFRESH", "Refresh"))

    CreateTable sheet, "T_LLLang", "M1:N3", _
                Array("CODE", "NAME"), Array(Array("EN", "English"), Array("FR", "Français"))

    CreateTable sheet, "T_SelectedLLLanguages", "P1:Q2", _
                Array("CODE", "NAME"), Array(Array("EN", "English"))
End Sub

Private Sub SetupNamedValues(ByVal sheet As Worksheet)
    AssignNamedValue sheet, "RNG_LLLanguage", "English"
    AssignNamedValue sheet, "RNG_LLLanguageCode", "EN"
    AssignNamedValue sheet, "RNG_GoToSection", "Section"
    AssignNamedValue sheet, "RNG_AnaPeriod", "Period"
    AssignNamedValue sheet, "RNG_GoToHeader", "Header"
    AssignNamedValue sheet, "RNG_DictionaryLanguage", "EN"
    AssignNamedValue sheet, "RNG_NoDevide", "NoDivide"
    AssignNamedValue sheet, "RNG_Devide", "Divide"
    AssignNamedValue sheet, "RNG_GoToGraph", "Graph"
    AssignNamedValue sheet, "RNG_OnFiltered", "Filtered"
    AssignNamedValue sheet, "RNG_CustomDrop", "Drop"
    AssignNamedValue sheet, "RNG_UASheet", "UA"
    AssignNamedValue sheet, "RNG_TSSheet", "TS"
    AssignNamedValue sheet, "RNG_SPSheet", "SP"
    AssignNamedValue sheet, "RNG_SPTSheet", "SPT"
    AssignNamedValue sheet, "RNG_CustomPivot", "Pivot"
    AssignNamedValue sheet, "RNG_Week", "Week"
    AssignNamedValue sheet, "RNG_Quarter", "Quarter"
    AssignNamedValue sheet, "RNG_InfoStart", "Start"
    AssignNamedValue sheet, "RNG_InfoEnd", "End"
    AssignNamedValue sheet, "RNG_SetupDictionarySheet", EXPORT_DICTIONARY_SHEET_NAME
End Sub

Private Sub AssignNamedValue(ByVal sheet As Worksheet, _
                             ByVal nameId As String, _
                             ByVal value As String)
    RemoveSheetName sheet, nameId
    sheet.Names.Add Name:=nameId, RefersTo:="=""" & Replace(value, """", """""") & """"
End Sub

Private Sub CreateTable(ByVal sheet As Worksheet, _
                        ByVal tableName As String, _
                        ByVal address As String, _
                        ByVal headers As Variant, _
                        ByVal rows As Variant)
    Dim headerIndex As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim targetRange As Range

    Set targetRange = sheet.Range(address)
    targetRange.Clear

    For headerIndex = LBound(headers) To UBound(headers)
        targetRange.Cells(1, headerIndex - LBound(headers) + 1).Value = headers(headerIndex)
    Next headerIndex

    For rowIndex = LBound(rows) To UBound(rows)
        For columnIndex = LBound(rows(rowIndex)) To UBound(rows(rowIndex))
            targetRange.Cells(rowIndex - LBound(rows) + 2, columnIndex - LBound(rows(rowIndex)) + 1).Value = rows(rowIndex)(columnIndex)
        Next columnIndex
    Next rowIndex

    sheet.ListObjects.Add xlSrcRange, targetRange, , xlYes
    sheet.ListObjects(sheet.ListObjects.Count).Name = tableName
End Sub

Private Sub PopulateDictionaryTable(ByVal sheet As Worksheet, _
                                    ByVal key As String, _
                                    ByVal translation As String)
    CreateTable sheet, DICTIONARY_TABLE_NAME(), "AA1:AC2", _
                Array("Key", "EN", "FR"), _
                Array(Array(key, translation, translation & " FR"))

    AssignDictionaryName sheet, sheet.ListObjects(DICTIONARY_TABLE_NAME())
End Sub

Private Function BuildWorkbookWithInlineDictionary(Optional ByVal dictionaryValue As String = "Hola") As Workbook
    Dim workbook As Workbook
    Set workbook = TestHelpers.NewWorkbook

    Dim translationSheet As Worksheet
    Set translationSheet = workbook.Worksheets(1)
    translationSheet.Name = TRANSLATION_SHEET_NAME

    SetupTranslationTables translationSheet
    SetupNamedValues translationSheet
    PopulateDictionaryTable translationSheet, "HELLO", dictionaryValue

    Set BuildWorkbookWithInlineDictionary = workbook
End Function

Private Function BuildWorkbookWithoutDictionary() As Workbook
    Dim workbook As Workbook
    Set workbook = TestHelpers.NewWorkbook

    Dim translationSheet As Worksheet
    Set translationSheet = workbook.Worksheets(1)
    translationSheet.Name = TRANSLATION_SHEET_NAME

    SetupTranslationTables translationSheet
    SetupNamedValues translationSheet

    Set BuildWorkbookWithoutDictionary = workbook
End Function

Private Function BuildSetupWorksheet(ByVal key As String, ByVal translation As String) As Worksheet
    Dim workbook As Workbook
    Set workbook = TestHelpers.NewWorkbook

    Dim sheet As Worksheet
    Set sheet = workbook.Worksheets(1)
    sheet.Name = "SetupDictionary"

    Dim rangeRef As Range
    Set rangeRef = sheet.Range("A1:C2")
    rangeRef.Cells(1, 1).Value = "Key"
    rangeRef.Cells(1, 2).Value = "EN"
    rangeRef.Cells(1, 3).Value = "FR"
    rangeRef.Cells(2, 1).Value = key
    rangeRef.Cells(2, 2).Value = translation
    rangeRef.Cells(2, 3).Value = translation & " FR"

    sheet.ListObjects.Add xlSrcRange, rangeRef, , xlYes
    sheet.ListObjects(1).Name = DICTIONARY_TABLE_NAME()
    AssignDictionaryName sheet, sheet.ListObjects(1)

    Set BuildSetupWorksheet = sheet
End Function

Private Function SheetHasName(ByVal sheet As Worksheet, ByVal name As String) As Boolean
    Dim idx As Long
    For idx = 1 To sheet.Names.Count
        If StrComp(ExtractSimpleName(sheet.Names(idx).Name), name, vbTextCompare) = 0 Then
            SheetHasName = True
            Exit Function
        End If
    Next idx
End Function

Private Function WorkbookHasSheet(ByVal book As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In book.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            WorkbookHasSheet = True
            Exit Function
        End If
    Next sheet
End Function

Private Sub RemoveSheetName(ByVal sheet As Worksheet, ByVal nameId As String)
    Dim idx As Long

    For idx = sheet.Names.Count To 1 Step -1
        If StrComp(ExtractSimpleName(sheet.Names(idx).Name), nameId, vbTextCompare) = 0 Then
            sheet.Names(idx).Delete
        End If
    Next idx
End Sub

Private Function SheetNameValue(ByVal sheet As Worksheet, ByVal nameId As String) As String
    Dim definition As Name
    Dim rng As Range

    For Each definition In sheet.Names
        If StrComp(ExtractSimpleName(definition.Name), nameId, vbTextCompare) = 0 Then
            On Error Resume Next
                Set rng = definition.RefersToRange
            On Error GoTo 0

            If Not rng Is Nothing Then
                SheetNameValue = CStr(rng.Value)
            Else
                SheetNameValue = DecodeNameDefinition(definition.RefersTo)
            End If
            Exit Function
        End If
    Next definition

    SheetNameValue = vbNullString
End Function

Private Function DecodeNameDefinition(ByVal refersTo As String) As String
    Dim text As String
    text = refersTo

    If Len(text) > 0 And Left$(text, 1) = "=" Then
        text = Mid$(text, 2)
    End If

    If Len(text) >= 2 Then
        If Left$(text, 1) = """" And Right$(text, 1) = """" Then
            text = Mid$(text, 2, Len(text) - 2)
        End If
    End If

    DecodeNameDefinition = Replace(text, """""", """")
End Function

Private Sub AssignDictionaryName(ByVal sheet As Worksheet, ByVal table As ListObject)
    RemoveSheetName sheet, "LLDictionaryTranslation"
    sheet.Names.Add Name:="LLDictionaryTranslation", RefersTo:=table.Range
End Sub

Private Function DictionaryTableIsEmpty(ByVal table As ListObject) As Boolean
    Dim data As Range

    Set data = table.DataBodyRange

    If data Is Nothing Then
        DictionaryTableIsEmpty = True
    Else
        DictionaryTableIsEmpty = (Application.WorksheetFunction.CountA(data) = 0)
    End If
End Function

Private Function ExtractSimpleName(ByVal qualifiedName As String) As String
    Dim exclPos As Long
    exclPos = InStr(qualifiedName, "!")
    If exclPos = 0 Then
        ExtractSimpleName = qualifiedName
    Else
        ExtractSimpleName = Mid$(qualifiedName, exclPos + 1)
    End If
End Function

Private Function DICTIONARY_TABLE_NAME() As String
    DICTIONARY_TABLE_NAME = "Tab_Translations"
End Function
