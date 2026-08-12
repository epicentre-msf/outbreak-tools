Attribute VB_Name = "TestDesignerEntry"
Attribute VB_Description = "Unit tests for DesignerEntry class"

Option Explicit

'@Folder("CustomTests.Designer")
'@ModuleDescription("Validates DesignerEntry for clearing, translation, AddInfo/ValueOf, TranslateMessage, Validate and sealing.")
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


'@TestMethod("DesignerEntry.Info")
Public Sub TestPasswordInfoNamesRoundTrip()
    CustomTestSetTitles Assert, "DesignerEntry", "TestPasswordInfoNamesRoundTrip"
    On Error GoTo Fail

    'Arrange: the two password ranges LinelistSpecs reads at generation.
    'The multi driver writes each row's passwords through these keys.
    FixtureWorkbook.Names.Add Name:="RNG_LLPwdOpen", RefersTo:=EntrySheet.Range("H1")
    FixtureWorkbook.Names.Add Name:="RNG_LLPassword", RefersTo:=EntrySheet.Range("H2")

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    subject.AddInfo "open-secret", "llpassword"
    subject.AddInfo "debug-secret", "debugpassword"

    'Assert: the values land on the ranges and read back by key
    Assert.AreEqual "open-secret", CStr(EntrySheet.Range("H1").value), _
                    "The llpassword key should write to RNG_LLPwdOpen."
    Assert.AreEqual "debug-secret", CStr(EntrySheet.Range("H2").value), _
                    "The debugpassword key should write to RNG_LLPassword."
    Assert.AreEqual "open-secret", subject.ValueOf("llpassword"), _
                    "ValueOf llpassword should read the written value back."
    Assert.AreEqual "debug-secret", subject.ValueOf("debugpassword"), _
                    "ValueOf debugpassword should read the written value back."
    Exit Sub

Fail:
    ReportTestFailure "TestPasswordInfoNamesRoundTrip"
End Sub

'@TestMethod("DesignerEntry.Info")
Public Sub TestEpiweekInfoNameRoundTrip()
    CustomTestSetTitles Assert, "DesignerEntry", "TestEpiweekInfoNameRoundTrip"
    On Error GoTo Fail

    'Arrange: the epiweek range LinelistSpecs reads at generation
    FixtureWorkbook.Names.Add Name:="RNG_DefaultEpiWeek", RefersTo:=EntrySheet.Range("H3")

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    'Act
    subject.AddInfo "2", "epiweekstart"

    'Assert
    Assert.AreEqual "2", CStr(EntrySheet.Range("H3").value), _
                    "The epiweekstart key should write to RNG_DefaultEpiWeek."
    Assert.AreEqual "2", subject.ValueOf("epiweekstart"), _
                    "ValueOf epiweekstart should read the written value back."
    Exit Sub

Fail:
    ReportTestFailure "TestEpiweekInfoNameRoundTrip"
End Sub


'@section Validate Tests
'===============================================================================
'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidatePassesOnValidEntries()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidatePassesOnValidEntries"
    On Error GoTo Fail

    ArrangeValidEntries

    Dim subject As DesignerEntry
    Set subject = DesignerEntry.Create(EntrySheet)

    Dim checks As Checking
    Set checks = subject.Validate()

    Assert.AreEqual CLng(0), CLng(checks.Length), _
                    "Validate should file no entry when every value is valid."
    Exit Sub

Fail:
    ReportTestFailure "TestValidatePassesOnValidEntries"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEmptySetupPath()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEmptySetupPath"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_PathDico").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("setup path"), _
                  "An empty setup path should file the setup path fault."
    Assert.IsTrue InStr(1, checks.ValueOf("setup path", checkingType), "Error", vbTextCompare) > 0, _
                  "The setup path fault should carry the error scope."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEmptySetupPath"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesUnfoundSetupFile()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesUnfoundSetupFile"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_PathDico").value = ThisWorkbook.Path & _
        Application.PathSeparator & "no_such_setup_file.xlsb"

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("setup path"), _
                  "A setup path absent from the disk should file the setup path fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesUnfoundSetupFile"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEmptyOutputFolder()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEmptyOutputFolder"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_LLDir").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("output folder"), _
                  "An empty output folder should file the output folder fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEmptyOutputFolder"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesUnfoundOutputFolder()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesUnfoundOutputFolder"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_LLDir").value = ThisWorkbook.Path & _
        Application.PathSeparator & "no_such_output_folder"

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("output folder"), _
                  "An output folder absent from the disk should file the output folder fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesUnfoundOutputFolder"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEmptyLinelistName()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEmptyLinelistName"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_LLName").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("linelist name"), _
                  "An empty linelist name should file the linelist name fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEmptyLinelistName"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEmptySetupLanguage()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEmptySetupLanguage"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_LangSetup").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("setup language"), _
                  "An empty setup language should file the setup language fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEmptySetupLanguage"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateRefusesTagShapedSetupLanguage()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateRefusesTagShapedSetupLanguage"
    On Error GoTo Fail

    'A tag column header can reach the language dropdown through the
    'fallback header read on a setup with no persisted language list.
    ArrangeValidEntries
    EntrySheet.Range("RNG_LangSetup").value = "__TagInternal__"

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("setup language"), _
                  "A tag-shaped setup language should file the setup language fault."
    Assert.IsTrue InStr(1, checks.ValueOf("setup language", checkingType), "Error", vbTextCompare) > 0, _
                  "The tag-shaped language fault should carry the error scope."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateRefusesTagShapedSetupLanguage"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEmptyLinelistLanguage()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEmptyLinelistLanguage"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_LLForm").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("linelist language"), _
                  "An empty linelist language should file the linelist language fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEmptyLinelistLanguage"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEmptyDesign()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEmptyDesign"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_DesignLL").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("design"), _
                  "An empty design value should file the design fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEmptyDesign"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateSkipsEmptyGeoPath()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateSkipsEmptyGeoPath"
    On Error GoTo Fail

    'The geobase is optional: an empty path is a valid entry
    ArrangeValidEntries
    EntrySheet.Range("RNG_PathGeo").value = vbNullString

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsFalse checks.KeyExists("geobase"), _
                   "An empty geobase path should file no fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateSkipsEmptyGeoPath"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesUnfoundGeoFile()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesUnfoundGeoFile"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_PathGeo").value = ThisWorkbook.Path & _
        Application.PathSeparator & "no_such_geobase.xlsx"

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("geobase"), _
                  "A geobase path absent from the disk should file the geobase fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesUnfoundGeoFile"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesUnfoundTemplateFile()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesUnfoundTemplateFile"
    On Error GoTo Fail

    ArrangeValidEntries
    EntrySheet.Range("RNG_LLTemp").value = ThisWorkbook.Path & _
        Application.PathSeparator & "no_such_template.xlsb"

    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.IsTrue checks.KeyExists("template"), _
                  "A template path absent from the disk should file the template fault."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesUnfoundTemplateFile"
End Sub

'@TestMethod("DesignerEntry.Validate")
Public Sub TestValidateFilesEveryFaultAtOnce()
    CustomTestSetTitles Assert, "DesignerEntry", "TestValidateFilesEveryFaultAtOnce"
    On Error GoTo Fail

    'A worksheet with none of the named ranges reads every value as
    'empty: the six required entries file, the two optional paths skip
    Dim checks As Checking
    Set checks = DesignerEntry.Create(EntrySheet).Validate()

    Assert.AreEqual CLng(6), CLng(checks.Length), _
                    "A bare worksheet should file the six required-entry faults."
    Exit Sub

Fail:
    ReportTestFailure "TestValidateFilesEveryFaultAtOnce"
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

'@sub-title Create the Main named ranges and fill them with valid entries
'@details
'The setup path points at this workbook and the output folder at its
'folder, so both disk checks pass. The geobase and template paths stay
'empty, the valid shape for the two optional entries.
Private Sub ArrangeValidEntries()
    FixtureWorkbook.Names.Add Name:="RNG_PathGeo", RefersTo:=EntrySheet.Range("A1")
    FixtureWorkbook.Names.Add Name:="RNG_PathDico", RefersTo:=EntrySheet.Range("A2")
    FixtureWorkbook.Names.Add Name:="RNG_LLName", RefersTo:=EntrySheet.Range("A3")
    FixtureWorkbook.Names.Add Name:="RNG_LLDir", RefersTo:=EntrySheet.Range("A4")
    FixtureWorkbook.Names.Add Name:="RNG_LLTemp", RefersTo:=EntrySheet.Range("A6")
    FixtureWorkbook.Names.Add Name:="RNG_LangSetup", RefersTo:=EntrySheet.Range("A7")
    FixtureWorkbook.Names.Add Name:="RNG_LLForm", RefersTo:=EntrySheet.Range("A8")
    FixtureWorkbook.Names.Add Name:="RNG_DesignLL", RefersTo:=EntrySheet.Range("A9")

    EntrySheet.Range("A2").value = ThisWorkbook.FullName
    EntrySheet.Range("A3").value = "my_linelist"
    EntrySheet.Range("A4").value = ThisWorkbook.Path
    EntrySheet.Range("A7").value = "ENG"
    EntrySheet.Range("A8").value = "FRA - Francais"
    EntrySheet.Range("A9").value = "Standard"
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
