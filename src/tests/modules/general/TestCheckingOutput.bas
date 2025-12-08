Attribute VB_Name = "TestCheckingOutput"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@TestModule
'@Folder("Tests")

Private Const DEFAULTCHECKINGSHEET As String = "CheckingOutputFixture"
Private Const DEFAULTFILTERCELL As String = "C1"
Private Const DEFAULTTITLEFILTERCELL As String = "E1"
Private Const DEFAULTTITLEDEFAULT As String = "All Titles"
Private Const DEFAULTTITLERANGE As String = "CheckingOutputTitles"
Private Const DEFAULTEVENTMAKER As String = "CheckingOutputEventInstalled"
Private Const DEFAULTROWMARKER As String = "CheckingOutputStartingRow"
Private Const HIDDEN_TITLE_COLUMN_INDEX As Long = 2
Private Const FIRST_VISIBLE_COLUMN_INDEX As Long = 3
Private Const FIRST_OUTPUT_ROW_INDEX As Long = 4
Private Const VISIBLE_COLUMN_COUNT As Long = 3

Private Assert As Object
Private Fakes As Object
Private OutputWriter As ICheckingOutput
Private PrimaryCheck As IChecking
Private SecondaryCheck As IChecking

'@section Helpers
'===============================================================================

Private Function BuildChecking(ByVal Heading As String, ByVal subHeading As String, _
                               ParamArray entries() As Variant) As IChecking
    Dim checkingInstance As IChecking
    Dim index As Long

    Set checkingInstance = checking.Create(Heading, subHeading)
    For index = LBound(entries) To UBound(entries)
        checkingInstance.Add entries(index)(0), entries(index)(1), entries(index)(2)
    Next index

    Set BuildChecking = checkingInstance
End Function

Private Function OutputSheet() As Worksheet
    Set OutputSheet = EnsureWorksheet(DEFAULTCHECKINGSHEET)
End Function

Private Function CountOccurrences(ByVal sh As Worksheet, ByVal textValue As String, _
                                  Optional ByVal includeHiddenColumns As Boolean = False) As Long
    Dim searchRange As Range
    Dim lastRow As Long

    If includeHiddenColumns Then
        Set searchRange = sh.UsedRange
    Else
        lastRow = sh.Cells(sh.Rows.Count, FIRST_VISIBLE_COLUMN_INDEX).End(xlUp).Row
        If lastRow < 1 Then lastRow = 1
        Set searchRange = sh.Range(sh.Cells(1, FIRST_VISIBLE_COLUMN_INDEX), _
                                   sh.Cells(lastRow, FIRST_VISIBLE_COLUMN_INDEX + VISIBLE_COLUMN_COUNT - 1))
    End If

    On Error Resume Next
        CountOccurrences = Application.WorksheetFunction.CountIf(searchRange, textValue)
    On Error GoTo 0
End Function

Private Function GetHiddenNameValue(ByVal sh As Worksheet, ByVal nameId As String) As String
    Dim store As IHiddenNames
    Set store = HiddenNames.Create(sh)
    GetHiddenNameValue = store.ValueAsString(nameId)
End Function

Private Function NormaliseTypeLabel(ByVal typeLabel As String) As String
    Dim cleaned As String

    cleaned = typeLabel
    cleaned = Replace(cleaned, ChrW(10060), vbNullString)
    cleaned = Replace(cleaned, ChrW(9888), vbNullString)
    cleaned = Replace(cleaned, ChrW(8505), vbNullString)
    cleaned = Replace(cleaned, ChrW(9998), vbNullString)
    cleaned = Replace(cleaned, ChrW(10004), vbNullString)

    NormaliseTypeLabel = Trim$(cleaned)
End Function

Private Function FindRowByHiddenTitle(ByVal sh As Worksheet, _
                                      ByVal targetTitle As String, _
                                      Optional ByVal requireDataLabel As Boolean = False) As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim storedTitle As String
    Dim hasLabel As Boolean

    lastRow = sh.Cells(sh.Rows.Count, FIRST_VISIBLE_COLUMN_INDEX).End(xlUp).Row

    For rowIndex = FIRST_OUTPUT_ROW_INDEX To lastRow
        storedTitle = Trim$(CStr(sh.Cells(rowIndex, HIDDEN_TITLE_COLUMN_INDEX).value))
        If StrComp(storedTitle, targetTitle, vbTextCompare) = 0 Then
            hasLabel = LenB(CStr(sh.Cells(rowIndex, FIRST_VISIBLE_COLUMN_INDEX + 1).value)) > 0
            If (requireDataLabel And hasLabel) Or (Not requireDataLabel) Then
                FindRowByHiddenTitle = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Sub ResetWorksheetModule(ByVal sh As Worksheet)
    Dim idx As Long
    Dim codeModule As Object
    Dim store As IHiddenNames

    On Error Resume Next
        For idx = sh.CustomProperties.Count To 1 Step -1
            If StrComp(sh.CustomProperties(idx).Name, DEFAULTEVENTMAKER, vbTextCompare) = 0 Then
                sh.CustomProperties(idx).Delete
                Exit For
            End If
        Next idx
        Set store = HiddenNames.Create(sh)
        store.RemoveName DEFAULTEVENTMAKER
        store.RemoveName DEFAULTROWMARKER
    On Error GoTo 0

    Set codeModule = sh.Parent.VBProject.VBComponents(sh.CodeName).CodeModule
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
End Sub

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet DEFAULTCHECKINGSHEET
    RestoreApp
    Application.EnableEvents = True
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    ClearWorksheet OutputSheet
    Set OutputWriter = CheckingOutput.Create(OutputSheet)
    Set PrimaryCheck = BuildChecking("Data Validation Summary", "Core checks", _
                                     Array("key-1", "Missing identifier", checkingError), _
                                     Array("key-2", "Inconsistent dates", checkingWarning))
    Set SecondaryCheck = BuildChecking("Data Validation Summary", "Extended checks", _
                                       Array("key-3", "Optional comment", checkingNote), _
                                       Array("key-4", "Informative remark", checkingInfo))
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set OutputWriter = Nothing
    Set PrimaryCheck = Nothing
    Set SecondaryCheck = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("CheckingOutput")
Private Sub TestPrintOutputWritesRows()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet
    Dim titleRange As Range
    Dim titleName As Name

    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.EnsureWorksheetChangeHandler
    OutputWriter.PrintOutput checks

    Set sh = OutputWriter.Wksh()

    Assert.AreEqual "Status:", sh.Cells(1, HIDDEN_TITLE_COLUMN_INDEX).value, "Status label should be present in column B"
    Assert.AreEqual "Title:", sh.Cells(1, sh.Range(DEFAULTTITLEFILTERCELL).Column - 1).value, "Title label should be placed before the dropdown"
    Assert.AreEqual "All", sh.Range(DEFAULTFILTERCELL).value, "Status filter should default to All"
    Assert.AreEqual DEFAULTTITLEDEFAULT, sh.Range(DEFAULTTITLEFILTERCELL).value, "Title filter should default to All Titles"
    Assert.AreEqual xlValidateList, sh.Range(DEFAULTFILTERCELL).Validation.Type, "Status cell should contain list validation"
    Assert.AreEqual xlValidateList, sh.Range(DEFAULTTITLEFILTERCELL).Validation.Type, "Title cell should contain list validation"
    Assert.IsTrue (GetHiddenNameValue(sh, DEFAULTEVENTMAKER) = "True"), "Worksheet change handler marker should be stored"
    Assert.IsTrue (CountOccurrences(sh, "Data Validation Summary") = 1), "Title should be written only once"
    Assert.AreEqual "Core checks", sh.Cells(7, FIRST_VISIBLE_COLUMN_INDEX).value, "Subtitle should follow the title"
    Assert.AreEqual "Error", NormaliseTypeLabel(sh.Cells(9, FIRST_VISIBLE_COLUMN_INDEX).value), "First data row should include type caption"
    Assert.AreEqual "Missing identifier", sh.Cells(9, FIRST_VISIBLE_COLUMN_INDEX + 1).value, "First data row should include label"
    Assert.IsTrue (RGB(192, 0, 0) = sh.Cells(9, FIRST_VISIBLE_COLUMN_INDEX + 1).Font.Color), "Error rows should use red font"
    Assert.AreEqual "End of checkings ", sh.Cells(sh.UsedRange.rows.Count, FIRST_VISIBLE_COLUMN_INDEX).value, "Last row should match final entry"
    Assert.IsTrue (RGB(112, 48, 160) = sh.Cells(14, FIRST_VISIBLE_COLUMN_INDEX).Font.Color), "Note rows should use purple font"
    Assert.IsTrue (RGB(244, 236, 255) = sh.Cells(14, FIRST_VISIBLE_COLUMN_INDEX).Interior.Color), "Note rows should use purple fill"
    Assert.AreEqual "Data Validation Summary", sh.Cells(9, HIDDEN_TITLE_COLUMN_INDEX).value, "Hidden title column should store parent title"
    Assert.IsTrue (RGB(255, 255, 255) = sh.Cells(9, HIDDEN_TITLE_COLUMN_INDEX).Font.Color), "Hidden title column should render with white font"

    On Error Resume Next
        Set titleName = sh.Names(DEFAULTTITLERANGE)
        Set titleRange = titleName.RefersToRange
    On Error GoTo Fail
    Assert.IsNotNothing titleName, "Named range object for titles should exist"
    Assert.AreEqual "= " & DEFAULTTITLERANGE, sh.Range(DEFAULTTITLEFILTERCELL).Validation.Formula1, "Title validation should reference the local named range"
    Assert.IsNotNothing titleRange, "Named range for titles should exist"
    Assert.IsTrue (titleRange.rows.Count = 2), "Title named range should include default and one title"
    Assert.AreEqual DEFAULTTITLEDEFAULT, CStr(titleRange.Cells(1, 1).value), "First item in title range should be default option"
    Assert.AreEqual "Data Validation Summary", CStr(titleRange.Cells(2, 1).value), "Second item in title range should match the written title"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestPrintOutputWritesRows"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestPrintOutputRejectsInvalidItems()
    Dim invalidChecks As BetterArray

    Set invalidChecks = BetterArrayFromList("invalid entry")
    On Error Resume Next
    OutputWriter.PrintOutput invalidChecks
    Assert.IsTrue (Err.Number = ProjectError.InvalidArgument), "PrintOutput should raise when items are not IChecking"
    On Error GoTo 0
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterDropdownHidesRows()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet

    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.EnsureWorksheetChangeHandler
    OutputWriter.PrintOutput checks

   Set sh = OutputWriter.Wksh()
    Application.EnableEvents = True

    Assert.AreEqual DEFAULTTITLEDEFAULT, sh.Range(DEFAULTTITLEFILTERCELL).value, "Title filter should default to All Titles"

    sh.Range(DEFAULTFILTERCELL).value = "Warnings"
    Assert.IsTrue sh.rows(9).Hidden, "Rows with non matching types should be hidden"
    Assert.IsFalse sh.rows(10).Hidden, "Rows with matching type should remain visible"
    Assert.IsFalse sh.rows(12).Hidden, "Subtitles should always be visible"

    sh.Range(DEFAULTFILTERCELL).value = "All"
    Assert.IsFalse sh.rows(9).Hidden, "All rows should be visible after resetting filter"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFilterDropdownHidesRows"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterWorksheetMethodAcceptsExplicitSelection()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet

    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.PrintOutput checks

   Set sh = OutputWriter.Wksh()
    Assert.AreEqual DEFAULTTITLEDEFAULT, sh.Range(DEFAULTTITLEFILTERCELL).value, "Title filter should remain at default before explicit filtering"
    OutputWriter.FilterWorksheet "Warnings"

    Assert.IsTrue sh.rows(9).Hidden, "Explicit warning filter should hide error rows"
    Assert.IsFalse sh.rows(10).Hidden, "Explicit warning filter should keep warning rows visible"
    Assert.IsTrue sh.rows(14).Hidden, "Explicit warning filter should hide note rows"
    Assert.IsTrue sh.rows(15).Hidden, "Explicit warning filter should hide info rows"
    Assert.IsFalse sh.rows(12).Hidden, "Section headers should remain visible"

    OutputWriter.FilterWorksheet "All"
    Assert.IsFalse sh.rows(9).Hidden, "All filter should reveal previously hidden rows"
    Assert.IsFalse sh.rows(10).Hidden, "All filter should keep warning rows visible"
    Assert.IsFalse sh.rows(14).Hidden, "All filter should restore note rows"
    Assert.IsFalse sh.rows(15).Hidden, "All filter should restore info rows"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFilterWorksheetMethodAcceptsExplicitSelection"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterWorksheetUsesCellValueWhenParameterMissing()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet
    Dim previousEventState As Boolean

    previousEventState = Application.EnableEvents
    Set checks = BetterArrayFromList(PrimaryCheck, SecondaryCheck)
    OutputWriter.PrintOutput checks

    Set sh = OutputWriter.Wksh()
    Application.EnableEvents = False
    sh.Range(DEFAULTFILTERCELL).value = "Notes"
    Application.EnableEvents = previousEventState

    OutputWriter.FilterWorksheet

    Assert.IsTrue sh.rows(9).Hidden, "Notes filter should hide error rows"
    Assert.IsTrue sh.rows(10).Hidden, "Notes filter should hide warning rows"
    Assert.IsFalse sh.rows(14).Hidden, "Notes filter should keep note rows visible"
    Assert.IsTrue sh.rows(15).Hidden, "Notes filter should hide info rows"
    Assert.IsFalse sh.rows(7).Hidden, "Section headers should remain visible when filtering by notes"
    Exit Sub

Fail:
    Application.EnableEvents = previousEventState
    FailUnexpectedError Assert, "TestFilterWorksheetUsesCellValueWhenParameterMissing"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestPrintOutputAppendsContent()
    On Error GoTo Fail

    Dim firstRun As BetterArray
    Dim secondRun As BetterArray
    Dim sh As Worksheet
    Dim firstEndRow As Long
    Dim secondEndRow As Long
    Dim titleRange As Range

    Set firstRun = BetterArrayFromList( _
        BuildChecking("Batch One", "First set", _
                      Array("key-a", "Original entry", checkingWarning)))
    OutputWriter.PrintOutput firstRun

    Set sh = OutputWriter.Wksh()
    firstEndRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).row

    Set secondRun = BetterArrayFromList( _
        BuildChecking("Batch Two", "Second set", _
                      Array("key-b", "Replacement entry", checkingInfo)))
    OutputWriter.PrintOutput secondRun

    secondEndRow = sh.Cells(sh.Rows.Count, 2).End(xlUp).row

    Assert.IsTrue (secondEndRow > firstEndRow), "Subsequent prints should append after existing content"
    Assert.IsTrue (CountOccurrences(sh, "Batch One") = 1), _
        "Existing titles should remain after subsequent prints"
    Assert.IsTrue (CountOccurrences(sh, "Batch Two") = 1), _
        "New title should be written once"
    Assert.IsTrue (CountOccurrences(sh, "Original entry") = 1), _
        "Existing entries should remain visible"
    Assert.IsTrue (CountOccurrences(sh, "Replacement entry") = 1), _
        "New entries should be present exactly once"

    On Error Resume Next
        Set titleRange = sh.Names(DEFAULTTITLERANGE).RefersToRange
    On Error GoTo Fail

    Assert.IsNotNothing titleRange, "Title named range should remain after appending"
    Assert.IsTrue (titleRange.rows.Count = 3), "Title named range should include default and two titles"
    Assert.AreEqual DEFAULTTITLEDEFAULT, CStr(titleRange.Cells(1, 1).value), "First item should stay the default option"
    Assert.AreEqual "Batch One", CStr(titleRange.Cells(2, 1).value), "Existing title should remain in the named range"
    Assert.AreEqual "Batch Two", CStr(titleRange.Cells(3, 1).value), "New title should extend the named range"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestPrintOutputAppendsContent"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestTitleFilterHasPriorityOverStatusFilter()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet
    Dim qualityCheck As IChecking
    Dim otherTitleRow As Long
    Dim qualityTitleRow As Long
    Dim qualityDataRow As Long

    Set qualityCheck = BuildChecking("Quality Checks", "Post-run checks", _
                                     Array("key-5", "Checklist review", checkingSuccess))

    Set checks = BetterArrayFromList(PrimaryCheck, qualityCheck)
    OutputWriter.PrintOutput checks

    Set sh = OutputWriter.Wksh()

    OutputWriter.FilterWorksheet , "Quality Checks"

    otherTitleRow = FindRowByHiddenTitle(sh, "Data Validation Summary")
    Assert.IsTrue (otherTitleRow > 0), "Primary title row should be located"
    Assert.IsTrue sh.rows(otherTitleRow).Hidden, "Filtering by title should hide unrelated titles"

    qualityTitleRow = FindRowByHiddenTitle(sh, "Quality Checks")
    Assert.IsTrue (qualityTitleRow > 0), "Quality Checks title row should be found"
    Assert.IsFalse sh.rows(qualityTitleRow).Hidden, "Requested title should remain visible"

    qualityDataRow = FindRowByHiddenTitle(sh, "Quality Checks", True)
    Assert.IsTrue (qualityDataRow > 0), "Quality Checks data row should exist"
    Assert.IsFalse sh.rows(qualityDataRow).Hidden, "Data rows should remain visible when only the title filter is applied"

    OutputWriter.FilterWorksheet "Errors", "Quality Checks"

    Assert.IsFalse sh.rows(qualityTitleRow).Hidden, "Title rows should stay visible under status filtering"
    Assert.IsTrue sh.rows(qualityDataRow).Hidden, "Rows without matching status should hide within the selected title"
    Assert.IsTrue sh.rows(otherTitleRow).Hidden, "Rows from other titles must stay hidden after applying both filters"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestTitleFilterHasPriorityOverStatusFilter"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestWorksheetChangeSyncsStatusWhenTitleChanges()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet
    Dim statusCell As Range
    Dim titleCell As Range
    Dim previousEventState As Boolean
    Dim qualityCheck As IChecking
    Dim qualityDataRow As Long
    Dim primaryDataRow As Long

    previousEventState = Application.EnableEvents

    Set qualityCheck = BuildChecking("Quality Checks", "Operations follow-up", _
                                     Array("key-q1", "Checklist outcome", checkingSuccess), _
                                     Array("key-q2", "Reminder", checkingInfo))

    Set checks = BetterArrayFromList(PrimaryCheck, qualityCheck)
    OutputWriter.EnsureWorksheetChangeHandler
    OutputWriter.PrintOutput checks

    Set sh = OutputWriter.Wksh()
    Set statusCell = sh.Range(DEFAULTFILTERCELL)
    Set titleCell = sh.Range(DEFAULTTITLEFILTERCELL)

    Application.EnableEvents = True

    statusCell.value = "Errors"
    Assert.AreEqual "Errors", CStr(statusCell.value), "Status drop-down should capture the requested severity"

    titleCell.value = "Quality Checks"

    Assert.AreEqual "All", CStr(statusCell.value), "Changing only the title filter should reset the status selection to All"

    qualityDataRow = FindRowByHiddenTitle(sh, "Quality Checks", True)
    Assert.IsTrue (qualityDataRow > 0), "Quality Checks detail row should be located"
    Assert.IsFalse sh.rows(qualityDataRow).Hidden, "Quality Checks entries should remain visible after the title filter update"

    primaryDataRow = FindRowByHiddenTitle(sh, "Data Validation Summary", True)
    Assert.IsTrue (primaryDataRow > 0), "Primary data row should exist"
    Assert.IsTrue sh.rows(primaryDataRow).Hidden, "Rows from other titles should hide when a specific title is chosen"

    Application.EnableEvents = previousEventState
    Exit Sub

Fail:
    Application.EnableEvents = previousEventState
    FailUnexpectedError Assert, "TestWorksheetChangeSyncsStatusWhenTitleChanges"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestFilterWorksheetIgnoresStatusWhenOnlyTitleProvided()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet
    Dim statusCell As Range
    Dim previousEventState As Boolean
    Dim qualityCheck As IChecking
    Dim otherTitleRow As Long
    Dim qualityDataRow As Long

    previousEventState = Application.EnableEvents

    Set qualityCheck = BuildChecking("Quality Checks", "Operations follow-up", _
                                     Array("key-q1", "Checklist outcome", checkingSuccess))

    Set checks = BetterArrayFromList(PrimaryCheck, qualityCheck)
    OutputWriter.PrintOutput checks

    Set sh = OutputWriter.Wksh()
    Set statusCell = sh.Range(DEFAULTFILTERCELL)

    Application.EnableEvents = False
    statusCell.value = "Errors"
    Application.EnableEvents = previousEventState

    OutputWriter.FilterWorksheet , "Quality Checks"

    qualityDataRow = FindRowByHiddenTitle(sh, "Quality Checks", True)
    Assert.IsTrue (qualityDataRow > 0), "Quality Checks data row should exist"
    Assert.IsFalse sh.rows(qualityDataRow).Hidden, "Data rows should remain visible when only the title filter is specified"

    otherTitleRow = FindRowByHiddenTitle(sh, "Data Validation Summary")
    Assert.IsTrue (otherTitleRow > 0), "Other title rows should exist"
    Assert.IsTrue sh.rows(otherTitleRow).Hidden, "Title filter must hide rows that belong to a different title"

    Exit Sub

Fail:
    Application.EnableEvents = previousEventState
    FailUnexpectedError Assert, "TestFilterWorksheetIgnoresStatusWhenOnlyTitleProvided"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestWorksheetChangeHandlerInjectsFilteringLogic()
    On Error GoTo Fail

    Dim checks As BetterArray
    Dim sh As Worksheet
    Dim moduleText As String
    Dim codeModule As Object
    Dim handlerCount As Long
    Dim handlerPosition As Long

    Set sh = OutputSheet()
    ResetWorksheetModule sh

    Set checks = BetterArrayFromList(PrimaryCheck)
    OutputWriter.PrintOutput checks
    OutputWriter.EnsureWorksheetChangeHandler

    Set codeModule = sh.Parent.VBProject.VBComponents(sh.CodeName).CodeModule
    If codeModule.CountOfLines > 0 Then
        moduleText = codeModule.Lines(1, codeModule.CountOfLines)
    End If

    Assert.IsFalse (LenB(moduleText) = 0), "Worksheet module should contain an injected handler"
    Assert.IsTrue (InStr(1, moduleText, "Private Sub Worksheet_Change", vbTextCompare) > 0), _
        "Injected code should declare Worksheet_Change locally"
    Assert.IsTrue (InStr(1, moduleText, "Private Sub FilterCheckingOutputRows", vbTextCompare) > 0), _
        "Filtering helper should be embedded alongside the handler"
    Assert.IsTrue (InStr(1, moduleText, "Private Function StripIcons", vbTextCompare) > 0), _
        "Icon stripping helper must be injected for worksheet execution"
    Assert.IsTrue (InStr(1, moduleText, "ResetUiState", vbTextCompare) > 0), _
        "Worksheet handler should restore application state"
    Assert.IsFalse (InStr(1, moduleText, "CheckingOutput.", vbTextCompare) > 0), _
        "Injected worksheet code must not depend on CheckingOutput members"

    handlerPosition = InStr(1, moduleText, "Private Sub Worksheet_Change", vbTextCompare)
    Do While handlerPosition > 0
        handlerCount = handlerCount + 1
        handlerPosition = InStr(handlerPosition + 1, moduleText, "Private Sub Worksheet_Change", vbTextCompare)
    Loop
    Assert.AreEqual 1, handlerCount, "Exactly one Worksheet_Change should be injected"

    ResetWorksheetModule sh
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestWorksheetChangeHandlerInjectsFilteringLogic"
End Sub

'@TestMethod("CheckingOutput")
Private Sub TestEnsureWorksheetChangeHandlerReplacesExistingCode()
    On Error GoTo Fail

    Dim sh As Worksheet
    Dim codeModule As Object
    Dim procStart As Long
    Dim procLines As Long
    Dim procText As String
    Dim firstCall As Long

    Set sh = OutputSheet()
    ResetWorksheetModule sh

    Set codeModule = sh.Parent.VBProject.VBComponents(sh.CodeName).CodeModule
    codeModule.InsertLines 1, "Private Sub Worksheet_Change(ByVal Target As Range)" & vbNewLine & _
                              "    On Error GoTo Handler" & vbNewLine & _
                              "    DummyChange Target" & vbNewLine & _
                              "ExitSub:" & vbNewLine & _
                              "    Exit Sub" & vbNewLine & _
                              "Handler:" & vbNewLine & _
                              "    Resume ExitSub" & vbNewLine & _
                              "End Sub" & vbNewLine & _
                              "Private Sub DummyChange(ByVal Target As Range)" & vbNewLine & _
                              "End Sub"

    OutputWriter.EnsureWorksheetChangeHandler
    OutputWriter.EnsureWorksheetChangeHandler

    procStart = codeModule.ProcStartLine("Worksheet_Change", 0)
    procLines = codeModule.ProcCountLines("Worksheet_Change", 0)
    procText = codeModule.Lines(procStart, procLines)

    Assert.IsTrue InStr(1, procText, "FilterCheckingOutputRows", vbTextCompare) > 0, _
        "Injected handler should call the embedded filtering routine"
    Assert.IsFalse InStr(1, procText, "DummyChange", vbTextCompare) > 0, _
        "Previous Worksheet_Change implementation must be replaced entirely"
    Assert.IsFalse InStr(1, procText, "CheckingOutput.", vbTextCompare) > 0, _
        "Injected worksheet code should not depend on CheckingOutput members"

    firstCall = InStr(1, procText, "Worksheet_Change", vbTextCompare)
    Assert.IsTrue InStr(firstCall + 1, procText, "Worksheet_Change", vbTextCompare) = 0, _
        "Duplicate Worksheet_Change procedures should not be present"

    ResetWorksheetModule sh
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestEnsureWorksheetChangeHandlerReplacesExistingCode"
End Sub
