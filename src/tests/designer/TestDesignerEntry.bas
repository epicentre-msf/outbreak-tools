Attribute VB_Name = "TestDesignerEntry"
Attribute VB_Description = "Unit tests for DesignerEntry class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates DesignerEntry for clearing, translation, AddInfo/ValueOf, TranslateMessage and sealing.")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As CustomTest
Private FixtureWorkbook As Workbook
Private EntrySheet As Worksheet
Private TranslationSheet As Worksheet

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"


'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDesignerEntry"
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
    Set EntrySheet = EnsureWorksheet("Main", FixtureWorkbook)
    Set TranslationSheet = EnsureWorksheet("DesignerTranslation", FixtureWorkbook)
End Sub

'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    On Error Resume Next
        DeleteWorkbook FixtureWorkbook
    On Error GoTo 0

    Set TranslationSheet = Nothing
    Set EntrySheet = Nothing
    Set FixtureWorkbook = Nothing

    RestoreApp
End Sub


'@section DesignerEntry Tests
'===============================================================================
'@TestMethod("DesignerEntry")
Public Sub TestClearResetsInputRanges()
    CustomTestSetTitles Assert, "DesignerEntry", "TestClearResetsInputRanges"
    On Error GoTo Fail

    'Arrange: create named ranges and set values + coloured backgrounds
    FixtureWorkbook.Names.Add Name:="RNG_PathGeo", RefersTo:=EntrySheet.Range("A1")
    FixtureWorkbook.Names.Add Name:="RNG_PathDico", RefersTo:=EntrySheet.Range("A2")
    FixtureWorkbook.Names.Add Name:="RNG_LLName", RefersTo:=EntrySheet.Range("A3")
    FixtureWorkbook.Names.Add Name:="RNG_LLDir", RefersTo:=EntrySheet.Range("A4")
    FixtureWorkbook.Names.Add Name:="RNG_Edition", RefersTo:=EntrySheet.Range("A5")
    FixtureWorkbook.Names.Add Name:="RNG_LLTemp", RefersTo:=EntrySheet.Range("A6")
    FixtureWorkbook.Names.Add Name:="RNG_LangSetup", RefersTo:=EntrySheet.Range("A7")
    FixtureWorkbook.Names.Add Name:="RNG_LLForm", RefersTo:=EntrySheet.Range("A8")

    EntrySheet.Range("A1").value = "/path/geo"
    EntrySheet.Range("A2").value = "/path/dico"
    EntrySheet.Range("A3").value = "my_linelist"
    EntrySheet.Range("A4").value = "/output"
    EntrySheet.Range("A5").value = "v2.1"
    EntrySheet.Range("A6").value = "/temp/path"
    EntrySheet.Range("A7").value = "ENG"
    EntrySheet.Range("A8").value = "form_value"

    Dim rng As Range
    For Each rng In EntrySheet.Range("A1:A8")
        rng.Interior.Color = vbYellow
    Next rng

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    subject.Clear

    'Assert: all values should be cleared and backgrounds set to white
    Dim idx As Long
    For idx = 1 To 8
        Assert.AreEqual vbNullString, CStr(EntrySheet.Cells(idx, 1).value), _
                        "Range A" & idx & " should be cleared."
        Assert.AreEqual CLng(vbWhite), CLng(EntrySheet.Cells(idx, 1).Interior.Color), _
                        "Range A" & idx & " background should be white."
    Next idx
    Exit Sub

Fail:
    ReportTestFailure "TestClearResetsInputRanges"
End Sub

'@TestMethod("DesignerEntry")
Public Sub TestClearResetsWarnings()
    CustomTestSetTitles Assert, "DesignerEntry", "TestClearResetsWarnings"
    On Error GoTo Fail

    'Arrange: create RNG_Warning and fill 3 contiguous warning cells
    FixtureWorkbook.Names.Add Name:="RNG_Warning", RefersTo:=EntrySheet.Range("B1")
    EntrySheet.Range("B1").value = "Warning 1"
    EntrySheet.Range("B2").value = "Warning 2"
    EntrySheet.Range("B3").value = "Warning 3"

    Dim rng As Range
    For Each rng In EntrySheet.Range("B1:B3")
        rng.Interior.Color = vbRed
    Next rng

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    subject.Clear

    'Assert: all 3 warning cells should be cleared
    Dim idx As Long
    For idx = 1 To 3
        Assert.AreEqual vbNullString, CStr(EntrySheet.Cells(idx, 2).value), _
                        "Warning cell B" & idx & " should be cleared."
        Assert.AreEqual CLng(vbWhite), CLng(EntrySheet.Cells(idx, 2).Interior.Color), _
                        "Warning cell B" & idx & " background should be white."
    Next idx
    Exit Sub

Fail:
    ReportTestFailure "TestClearResetsWarnings"
End Sub

'@TestMethod("DesignerEntry")
Public Sub TestTranslateUpdatesLanguageCode()
    CustomTestSetTitles Assert, "DesignerEntry", "TestTranslateUpdatesLanguageCode"
    On Error GoTo Fail

    Dim subject As DesignerEntry
    Dim translator As DesignerTranslation

    Set translator = MakeDesignerTranslator()
    FixtureWorkbook.Names.Add Name:="RNG_DesignerTitle", RefersTo:=EntrySheet.Range("A1")

    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator
    subject.Translate "ENG"

    Assert.IsTrue Not (translator.TransObject Is Nothing), _
                  "Translate should set the designer language so translation objects resolve."
    Assert.AreEqual "Designer Title", CStr(EntrySheet.Range("RNG_DesignerTitle").value), _
                    "Translate should apply the designer range translation to the entry sheet."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateUpdatesLanguageCode"
End Sub


'@section DesignerEntry AddInfo/ValueOf Tests
'===============================================================================
'@TestMethod("DesignerEntry.Info")
Public Sub TestAddInfoWritesToNamedRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestAddInfoWritesToNamedRange"
    On Error GoTo Fail

    'Arrange: create the setuppath named range on the entry sheet
    FixtureWorkbook.Names.Add Name:="RNG_PathDico", RefersTo:=EntrySheet.Range("B1")

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    subject.AddInfo "/path/to/setup.xlsb", "setuppath"

    'Assert
    Assert.AreEqual "/path/to/setup.xlsb", CStr(EntrySheet.Range("B1").value), _
                    "AddInfo should write the value to the named range."
    Assert.AreEqual CLng(vbWhite), CLng(EntrySheet.Range("B1").Interior.Color), _
                    "AddInfo should set the cell background to white."
    Exit Sub

Fail:
    ReportTestFailure "TestAddInfoWritesToNamedRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestAddInfoEditionRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestAddInfoEditionRange"
    On Error GoTo Fail

    'Arrange
    FixtureWorkbook.Names.Add Name:="RNG_Edition", RefersTo:=EntrySheet.Range("C1")

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: write a status message to the edition range
    subject.AddInfo "File loaded", "edition"

    'Assert
    Assert.AreEqual "File loaded", CStr(EntrySheet.Range("C1").value), _
                    "Edition range should contain the message."
    Exit Sub

Fail:
    ReportTestFailure "TestAddInfoEditionRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfReadsFromNamedRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfReadsFromNamedRange"
    On Error GoTo Fail

    'Arrange: create the geopath named range and write a value
    FixtureWorkbook.Names.Add Name:="RNG_PathGeo", RefersTo:=EntrySheet.Range("D1")
    EntrySheet.Range("D1").value = "/path/to/geo.xlsx"

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("geopath")

    'Assert
    Assert.AreEqual "/path/to/geo.xlsx", result, _
                    "ValueOf should return the value from the named range."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfReadsFromNamedRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfReturnsEmptyForUnknownRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfReturnsEmptyForUnknownRange"
    On Error GoTo Fail

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: request an unknown info name
    Dim result As String
    result = subject.ValueOf("nonexistent")

    'Assert
    Assert.AreEqual vbNullString, result, _
                    "ValueOf should return empty for unknown info names."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfReturnsEmptyForUnknownRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestTranslateMessageReturnsTranslatedText()
    CustomTestSetTitles Assert, "DesignerEntry", "TestTranslateMessageReturnsTranslatedText"
    On Error GoTo Fail

    'Arrange: the fixture seeds T_tradMsg with MSG_ChemFich -> "File path loaded"
    Dim translator As DesignerTranslation
    Set translator = MakeDesignerTranslator()

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator
    subject.Translate "ENG"    'set the designer language so message lookups resolve

    'Act
    Dim result As String
    result = subject.TranslateMessage("MSG_ChemFich")

    'Assert
    Assert.AreEqual "File path loaded", result, _
                    "TranslateMessage should return the translated text."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateMessageReturnsTranslatedText"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestTranslateMessageFallsBackToRawCode()
    CustomTestSetTitles Assert, "DesignerEntry", "TestTranslateMessageFallsBackToRawCode"
    On Error GoTo Fail

    'Arrange: a real translator whose message table has no MSG_Unknown entry
    Dim translator As DesignerTranslation
    Set translator = MakeDesignerTranslator()

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator translator
    subject.Translate "ENG"

    'Act
    Dim result As String
    result = subject.TranslateMessage("MSG_Unknown")

    'Assert: unknown codes fall back to the raw message code
    Assert.AreEqual "MSG_Unknown", result, _
                    "TranslateMessage should fall back to the raw message code."
    Exit Sub

Fail:
    ReportTestFailure "TestTranslateMessageFallsBackToRawCode"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestAddInfoSilentlySkipsMissingRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestAddInfoSilentlySkipsMissingRange"
    On Error GoTo Fail

    'Arrange: do NOT create the RNG_PathDico named range

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: should not raise an error
    subject.AddInfo "/some/path", "setuppath"

    'Assert: no error was raised (test completes)
    Assert.IsTrue True, "AddInfo should silently skip when the named range is missing."
    Exit Sub

Fail:
    ReportTestFailure "TestAddInfoSilentlySkipsMissingRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfLLDirRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfLLDirRange"
    On Error GoTo Fail

    'Arrange: create the lldir named range and write a value
    FixtureWorkbook.Names.Add Name:="RNG_LLDir", RefersTo:=EntrySheet.Range("E1")
    EntrySheet.Range("E1").value = "/output/folder"

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("lldir")

    'Assert
    Assert.AreEqual "/output/folder", result, _
                    "ValueOf lldir should return the linelist directory path."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfLLDirRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfLLNameRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfLLNameRange"
    On Error GoTo Fail

    'Arrange
    FixtureWorkbook.Names.Add Name:="RNG_LLName", RefersTo:=EntrySheet.Range("F1")
    EntrySheet.Range("F1").value = "my_linelist"

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("llname")

    'Assert
    Assert.AreEqual "my_linelist", result, _
                    "ValueOf llname should return the linelist name."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfLLNameRange"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestValueOfTempPathRange()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValueOfTempPathRange"
    On Error GoTo Fail

    'Arrange
    FixtureWorkbook.Names.Add Name:="RNG_LLTemp", RefersTo:=EntrySheet.Range("G1")
    EntrySheet.Range("G1").value = "/path/to/template.xlsb"

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    Dim result As String
    result = subject.ValueOf("temppath")

    'Assert
    Assert.AreEqual "/path/to/template.xlsb", result, _
                    "ValueOf temppath should return the template file path."
    Exit Sub

Fail:
    ReportTestFailure "TestValueOfTempPathRange"
End Sub


'@section Sealing Tests
'===============================================================================
'@TestMethod("DesignerEntry.Sealing")
Public Sub TestConfigureIsSetOnce()
    CustomTestSetTitles Assert, "DesignerEntry", "TestConfigureIsSetOnce"
    On Error GoTo Fail

    'Arrange: Create configures and seals the instance
    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: a second Configure should raise
    Dim raisedNumber As Long
    On Error Resume Next
    subject.Configure EntrySheet
    raisedNumber = Err.Number
    On Error GoTo Fail

    'Assert
    Assert.IsTrue raisedNumber <> 0, "Configure should raise on a sealed instance."
    Assert.IsNotNothing subject.HostSheet, "The first binding should survive the refused call."
    Exit Sub

Fail:
    ReportTestFailure "TestConfigureIsSetOnce"
End Sub

'@TestMethod("DesignerEntry.Sealing")
Public Sub TestUseTranslatorRefusesNothing()
    CustomTestSetTitles Assert, "DesignerEntry", "TestUseTranslatorRefusesNothing"
    On Error GoTo Fail

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act: injecting Nothing should raise
    Dim raisedNumber As Long
    On Error Resume Next
    subject.UseTranslator Nothing
    raisedNumber = Err.Number
    On Error GoTo Fail

    'Assert
    Assert.IsTrue raisedNumber <> 0, "UseTranslator should refuse Nothing."
    Exit Sub

Fail:
    ReportTestFailure "TestUseTranslatorRefusesNothing"
End Sub

'@TestMethod("DesignerEntry.Sealing")
Public Sub TestUseTranslatorIsSetOnce()
    CustomTestSetTitles Assert, "DesignerEntry", "TestUseTranslatorIsSetOnce"
    On Error GoTo Fail

    'Arrange: a translator is already held
    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)
    subject.UseTranslator MakeDesignerTranslator()

    'Act: a second injection should raise
    Dim raisedNumber As Long
    On Error Resume Next
    subject.UseTranslator MakeDesignerTranslator()
    raisedNumber = Err.Number
    On Error GoTo Fail

    'Assert
    Assert.IsTrue raisedNumber <> 0, "UseTranslator should refuse to replace a held translator."
    Exit Sub

Fail:
    ReportTestFailure "TestUseTranslatorIsSetOnce"
End Sub


'@section Internal helpers
'===============================================================================

Private Sub ReportTestFailure(ByVal context As String)
    Dim message As String

    If Assert Is Nothing Then Exit Sub

    message = context & " failed with error " & Err.Number & " (" & Err.Source & "): " & Err.Description
    Assert.LogFailure message
    Err.Clear
End Sub

'@sub-title Build a real DesignerTranslation backed by a seeded translation sheet
Private Function MakeDesignerTranslator() As DesignerTranslation
    Dim sh As Worksheet
    Set sh = EnsureWorksheet("DesignerTradTables", FixtureWorkbook)
    SeedDesignerTranslationTables sh
    Set MakeDesignerTranslator = DesignerTranslation.Create(sh)
End Function

'@sub-title Seed the four designer translation ListObjects (ENG) required by DesignerTranslation
Private Sub SeedDesignerTranslationTables(ByVal sh As Worksheet)
    sh.Cells.Clear
    AddTradTable sh, sh.Range("A1"), "T_tradMsg", _
        Array(Array("tag", "ENG"), Array("MSG_ChemFich", "File path loaded"), Array("MSG_Info", "Information"))
    AddTradTable sh, sh.Range("D1"), "T_tradShape", _
        Array(Array("tag", "ENG"), Array("shp_title", "Designer"))
    AddTradTable sh, sh.Range("G1"), "T_tradRange", _
        Array(Array("tag", "ENG"), Array("RNG_DesignerTitle", "Designer Title"))
    AddTradTable sh, sh.Range("J1"), "T_tradDrop", _
        Array(Array("tag", "ENG"), Array("drp_choice", "list_values"))
End Sub

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
