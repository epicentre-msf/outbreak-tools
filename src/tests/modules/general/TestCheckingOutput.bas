Attribute VB_Name = "TestCheckingOutput"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the CheckingOutput class")

'@description
'   Tests the CheckingOutput class which writes IChecking entries to a worksheet
'   and provides filter dropdowns for status and title. Tests cover PrintOutput row
'   writing, filter dropdown validation, explicit and cell-based filtering,
'   title-priority filtering when combined with status, Worksheet_Change handler
'   injection, handler replacement of existing code, and append behavior across
'   multiple print calls.
'@depends CheckingOutput, ICheckingOutput, Checking, IChecking, HiddenNames, BetterArray, TestHelpers

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

'@sub-title Build an IChecking instance from a heading, subheading, and entry arrays
'@details Each element of entries is expected to be a three-element Array(key, label,
'   checkingType). The function creates a Checking, adds all entries, and returns
'   the resulting IChecking interface.
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

'@sub-title Return the fixture worksheet, creating it if absent
Private Function OutputSheet() As Worksheet
    Set OutputSheet = EnsureWorksheet(DEFAULTCHECKINGSHEET)
End Function

'@sub-title Count visible-column occurrences of a text value on a worksheet
'@details When includeHiddenColumns is False (default) the search is limited to
'   columns FIRST_VISIBLE_COLUMN_INDEX through FIRST_VISIBLE_COLUMN_INDEX +
'   VISIBLE_COLUMN_COUNT - 1. When True, the entire UsedRange is searched.
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

'@sub-title Read a worksheet-level hidden name as a string
Private Function GetHiddenNameValue(ByVal sh As Worksheet, ByVal nameId As String) As String
    Dim store As IHiddenNames
    Set store = HiddenNames.Create(sh)
    GetHiddenNameValue = store.ValueAsString(nameId)
End Function

'@sub-title Strip emoji icons from a type label and return the plain text
'@details Removes cross mark, warning, info, scissors, and check mark Unicode
'   characters that CheckingOutput prepends to type labels so that assertions
'   can compare the underlying English text.
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

'@sub-title Locate the first output row whose hidden title column matches a target
'@details Scans from FIRST_OUTPUT_ROW_INDEX downward in column HIDDEN_TITLE_COLUMN_INDEX.
'   When requireDataLabel is True, the row must also have a non-empty data label in
'   the column after the first visible column, which distinguishes data rows from
'   title/subtitle rows.
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

'@sub-title Remove the event-installed marker and clear the worksheet code module
'@details Deletes the CustomProperty and HiddenNames entries for the event and row
'   markers, then removes all lines from the worksheet's VBComponent code module
'   so that subsequent tests start with a clean slate.
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

'@sub-title Verify that PrintOutput writes rows, headers, dropdowns, colours, and named ranges
'@details Arranges two IChecking objects with error, warning, note, and info entries,
'   calls PrintOutput, then asserts on cell values (title, subtitle, type labels, data
'   labels), dropdown validations, font and fill colours per type, the hidden title
'   column content and styling, the event-installed marker, and the title named range
'   contents and size.
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
    Assert.IsTrue (RGB(112, 48, 160) = sh.Cells(16, FIRST_VISIBLE_COLUMN_INDEX).Font.Color), "Note rows should use purple font"
    Assert.IsTrue (RGB(244, 236, 255) = sh.Cells(16, FIRST_VISIBLE_COLUMN_INDEX).Interior.Color), "Note rows should use purple fill"
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

'@sub-title Verify that PrintOutput raises an error when items are not IChecking
'@details Arranges a BetterArray containing a plain string instead of an IChecking
'   object, calls PrintOutput, and asserts that the resulting error number equals
'   ProjectError.InvalidArgument.
'@TestMethod("CheckingOutput")
Private Sub TestPrintOutputRejectsInvalidItems()
    Dim invalidChecks As BetterArray

    Set invalidChecks = BetterArrayFromList("invalid entry")
    On Error Resume Next
    OutputWriter.PrintOutput invalidChecks
    Assert.IsTrue (Err.Number = ProjectError.InvalidArgument), "PrintOutput should raise when items are not IChecking"
    On Error GoTo 0
End Sub

'@sub-title Verify that the status dropdown hides and reveals rows via Worksheet_Change
'@details Arranges output with all four checking types, enables events, sets the
'   status filter cell to "Warnings", and asserts that non-warning data rows are
'   hidden while warning rows and subtitles remain visible. Resets the filter to
'   "All" and asserts that all rows are restored.
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

'@sub-title Verify that FilterWorksheet accepts an explicit status parameter
'@details Arranges output with all four types, calls FilterWorksheet("Warnings"), and
'   asserts that only warning rows and section headers remain visible. Then calls
'   FilterWorksheet("All") and asserts all rows are restored to visible.
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

'@sub-title Verify that FilterWorksheet reads the filter cell value when no parameter is given
'@details Arranges output, programmatically sets the status filter cell to "Notes"
'   while events are disabled, then calls FilterWorksheet with no arguments. Asserts
'   that only note rows and section headers remain visible, confirming the method
'   falls back to the cell value.
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

'@sub-title Verify that PrintOutput appends content and extends the title named range
'@details Writes a first IChecking ("Batch One"), records the last row, then writes a
'   second IChecking ("Batch Two") and asserts the last row increased. Confirms that
'   both titles and their entries appear exactly once, and that the title named range
'   now contains three entries: the default, Batch One, and Batch Two.
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

'@sub-title Verify that the title filter takes priority over the status filter
'@details Arranges two IChecking groups under different titles, applies a title-only
'   filter for "Quality Checks", then asserts that rows from the other title are
'   hidden while all Quality Checks rows remain visible. Next applies both a status
'   filter ("Errors") and the title filter, asserting that non-matching status rows
'   within the selected title hide while the title row stays visible and rows from
'   other titles remain hidden.
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

'@sub-title Verify that changing the title filter resets the status filter to "All"
'@details Arranges two titled groups and enables Worksheet_Change events. Sets the
'   status filter to "Errors", then changes the title filter to "Quality Checks".
'   Asserts that the status cell resets to "All", Quality Checks data rows remain
'   visible, and rows from the primary title are hidden by the title constraint.
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

'@sub-title Verify that FilterWorksheet ignores the status cell when only a title is provided
'@details Arranges output with two titled groups, sets the status cell to "Errors"
'   while events are disabled, then calls FilterWorksheet with only the title
'   parameter ("Quality Checks"). Asserts that all Quality Checks data rows remain
'   visible regardless of the lingering status value, and that rows from the other
'   title are hidden.
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

'@sub-title Verify that EnsureWorksheetChangeHandler injects a self-contained filtering handler
'@details Resets the worksheet code module, calls PrintOutput and
'   EnsureWorksheetChangeHandler, then reads the full module text. Asserts that the
'   module contains exactly one Worksheet_Change declaration, a
'   FilterCheckingOutputRows helper, a StripIcons function, a ResetUiState call,
'   and no external references to CheckingOutput members.
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

'@sub-title Verify that EnsureWorksheetChangeHandler replaces an existing Worksheet_Change
'@details Seeds the worksheet code module with a dummy Worksheet_Change and a companion
'   DummyChange sub, then calls EnsureWorksheetChangeHandler twice. Asserts that the
'   resulting handler calls FilterCheckingOutputRows, no longer references DummyChange,
'   contains no external CheckingOutput dependencies, and that only one
'   Worksheet_Change declaration exists.
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
