Attribute VB_Name = "EventsDesignerMulti"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the Multi group on the designer workbook, and the driver that generates one linelist per row.")
'@depends CustomTable, ApplicationState, OSFiles, DropdownLists, BetterArray, EventsDesignerAdvanced, EventsDesignerCore, DesignerEntry, GenerationReport, Checking, ProgressBar, Linelist
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Ribbon callbacks for the Multi group manage the T_Multi ListObject on
'the GenerateMultiple worksheet. Each callback follows the established
'pattern: show dialogs before entering busy state, wrap work in
'On Error GoTo Cleanup, and restore application state on exit.
'
'THE ID RULE
'-------------------------------------------------------------------------------
'An ID is written once. Adding rows, duplicating a row, resizing and
'importing all fill only the blank ID cells with the next free numbers,
'through EnsureRowIds. The per-row language dropdown is named after the
'row ID (<id>_lang), so a rewritten ID would detach its row from the
'dropdown the validation points at.
'
'A step that gets skipped is reported: the callbacks collect one line
'per skipped step and show them in one message after the busy state is
'restored.
'
'THE MULTI GENERATION
'-------------------------------------------------------------------------------
'clickGenerateMulti walks the table and runs the single-build core
'(EventsDesignerAdvanced.GenerateOne) once per filled row. The row's
'values land on Main through the shared DesignerEntry, so the build
'reads them from the same ranges the single generation reads; the setup
'workbook of a row is opened once, inside the build. Every row flushes
'its checkings into the one generation report of the run, its outcome
'lands in the result column, and a failed row leaves the loop running.

Private Const SHEET_GENERATE_MULTIPLE As String = "GenerateMultiple"
Private Const TABLE_MULTI As String = "T_Multi"
Private Const PROMPT_TITLE As String = "Designer"

'The twelve columns of T_Multi. The callbacks here write the three path
'columns and wire the dictionary language dropdown; the driver reads the
'row into the Main entries and writes result and output files back.
Private Const COL_ID As String = "ID"
Private Const COL_SETUPS As String = "setups"
Private Const COL_GEOBASES As String = "geobases"
Private Const COL_OUTPUT_FOLDERS As String = "output folders"
Private Const COL_OUTPUT_FILES As String = "output files"
Private Const COL_PASSWORD As String = "output file password"
Private Const COL_DEBUG_PASSWORD As String = "output file debugging password"
Private Const COL_LANG_DICTIONARY As String = "language of the dictionary"
Private Const COL_LANG_INTERFACE As String = "language of the interface"
Private Const COL_EPIWEEK_START As String = "epiweek start"
Private Const COL_DESIGN As String = "design"
Private Const COL_RESULT As String = "result"

'Setup language extraction
Private Const SHEET_TRANSLATIONS As String = "Translations"
Private Const ID_PREFIX As String = "Operation-"

'Dropdown-based language validation
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const DROPDOWN_PREFIX As String = "dropdown_"
Private Const LANG_SUFFIX As String = "_lang"

Private Const MSG_PLACE_DATA As String = "Please place the cursor inside the table data area."

'Progress bar over the rows. The two names live on the GenerateMultiple
'sheet and are owner hand work on the mock; a missing name means no bar
'and no raise, so this code lands before the hand work and wakes up
'with it.
Private Const RNG_PROGRESS_BAR As String = "RNG_ProgressBar"
Private Const RNG_PROGRESS_STATUS As String = "RNG_ProgressStatus"

'The result column value of a row that built and saved
Private Const RESULT_BUILT As String = "OK"


'@section Multi group callbacks
'===============================================================================

'@Description("Load files or folder into the active T_Multi column (setups, geobases, output folders).")
'@EntryPoint
Public Sub clickFolderMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim colName As String
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim skipped As BetterArray
    Dim startRow As Long

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    'The write lands at the cursor row, so the cursor has to sit on a
    'data cell. A cursor on the header row used to write the first file
    'path over the header text.
    If lo.DataBodyRange Is Nothing Then
        MsgBox MSG_PLACE_DATA, vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If
    If Intersect(Application.ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox MSG_PLACE_DATA, vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    colName = ActiveCellColumnName(lo)

    'Validate that the active cell is on a supported column
    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS), LCase$(COL_GEOBASES), LCase$(COL_OUTPUT_FOLDERS)
        'valid column, continue
    Case Else
        MsgBox "Please place the cursor on the " & Chr(34) & COL_SETUPS & Chr(34) & _
               ", " & Chr(34) & COL_GEOBASES & Chr(34) & ", or " & Chr(34) & _
               COL_OUTPUT_FOLDERS & Chr(34) & " column.", _
               vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End Select

    'Show the appropriate file/folder dialog before entering busy state
    Set io = OSFiles.Create()

    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS)
        io.LoadFiles "*.xlsb;*.xlsx"
        If Not io.HasValidFiles() Then Exit Sub
    Case LCase$(COL_GEOBASES)
        io.LoadFiles "*.xlsx"
        If Not io.HasValidFiles() Then Exit Sub
    Case LCase$(COL_OUTPUT_FOLDERS)
        io.LoadFolder
        If Not io.HasValidFolder() Then Exit Sub
    End Select

    startRow = Application.ActiveCell.Row

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set skipped = New BetterArray
    skipped.LowerBound = 1

    Select Case LCase$(colName)
    Case LCase$(COL_SETUPS)
        LoadSetupFiles lo, CollectFiles(io), startRow, ResolveDropdownManager(), skipped
    Case LCase$(COL_GEOBASES)
        LoadGeobaseFiles lo, CollectFiles(io), startRow, skipped
    Case LCase$(COL_OUTPUT_FOLDERS)
        LoadOutputFolder lo, io.Folder(), startRow, skipped
    End Select

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickFolderMulti: "; errNumber; errDesc
        MsgBox "Unable to load files: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not skipped Is Nothing Then
        ShowSkipped skipped
    End If
End Sub

'@Description("Duplicate the active row in T_Multi with the same values and a fresh ID.")
'@EntryPoint
Public Sub clickDupMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim appScope As ApplicationState
    Dim relPos As Long
    Dim sourceRow As Range
    Dim destRow As Range
    Dim idCol As ListColumn
    Dim idMissing As Boolean

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    'Verify the active cell is inside the table data body
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Application.ActiveCell, lo.DataBodyRange) Is Nothing Then
        MsgBox MSG_PLACE_DATA, vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    'Compute the relative row position (1-based within ListRows)
    relPos = Application.ActiveCell.Row - lo.HeaderRowRange.Row

    'Insert a new row immediately below the current one
    If relPos >= lo.ListRows.Count Then
        lo.ListRows.Add
    Else
        lo.ListRows.Add Position:=relPos + 1
    End If

    'Copy values from the source row to the new row
    Set sourceRow = lo.ListRows(relPos).Range
    Set destRow = lo.ListRows(relPos + 1).Range
    destRow.Value = sourceRow.Value

    'The copy carried the source row's ID, and two rows sharing an ID
    'share the <id>_lang dropdown. The new row starts blank and gets the
    'next free number.
    On Error Resume Next
    Set idCol = lo.ListColumns(COL_ID)
    On Error GoTo Cleanup

    If idCol Is Nothing Then
        idMissing = True
    Else
        destRow.Cells(1, idCol.Index).Value = vbNullString
        EnsureRowIds lo
    End If

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickDupMulti: "; errNumber; errDesc
        MsgBox "Unable to duplicate row: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf idMissing Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Add rows to the T_Multi table. New rows get the next free IDs.")
'@EntryPoint
Public Sub clickAddRowsMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim table As CustomTable
    Dim appScope As ApplicationState
    Dim hasIdColumn As Boolean

    hasIdColumn = True

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set table = CustomTable.Create(lo)
    table.AddRows nbRows:=10, insertShift:=False, includeIds:=False
    hasIdColumn = EnsureRowIds(lo)

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickAddRowsMulti: "; errNumber; errDesc
        MsgBox "Unable to add rows: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not hasIdColumn Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Resize the T_Multi table by removing empty rows. Kept rows keep their IDs.")
'@EntryPoint
Public Sub clickResizeMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim table As CustomTable
    Dim appScope As ApplicationState
    Dim hasIdColumn As Boolean

    hasIdColumn = True

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set table = CustomTable.Create(lo)
    table.RemoveRows totalCount:=0, includeIds:=False, forceShift:=False
    hasIdColumn = EnsureRowIds(lo)

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickResizeMulti: "; errNumber; errDesc
        MsgBox "Unable to resize table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not hasIdColumn Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Import T_Multi data from another workbook. Blank IDs get the next free numbers.")
'@EntryPoint
Public Sub clickImpMulti(ByRef control As IRibbonControl)
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim importBook As Workbook
    Dim sourceLo As ListObject
    Dim targetLo As ListObject
    Dim sourceTable As CustomTable
    Dim targetTable As CustomTable
    Dim hasIdColumn As Boolean

    hasIdColumn = True

    'Show file picker before entering busy state
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb;*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=True)

    'Try to find T_Multi on the GenerateMultiple sheet first
    On Error Resume Next
    Set sourceLo = importBook.Worksheets(SHEET_GENERATE_MULTIPLE).ListObjects(TABLE_MULTI)
    On Error GoTo Cleanup

    'Fallback: use the first ListObject on the first worksheet
    If sourceLo Is Nothing Then
        If importBook.Worksheets(1).ListObjects.Count > 0 Then
            Set sourceLo = importBook.Worksheets(1).ListObjects(1)
        End If
    End If

    If sourceLo Is Nothing Then
        importBook.Close saveChanges:=False
        Set importBook = Nothing
        MsgBox "No table found in the selected workbook.", _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        GoTo Cleanup
    End If

    Set targetLo = ResolveMultiTable()
    If targetLo Is Nothing Then
        importBook.Close saveChanges:=False
        Set importBook = Nothing
        ReportMissingTable
        GoTo Cleanup
    End If

    Set sourceTable = CustomTable.Create(sourceLo)
    Set targetTable = CustomTable.Create(targetLo)
    targetTable.Import sourceTable
    hasIdColumn = EnsureRowIds(targetLo)

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not importBook Is Nothing Then
        importBook.Close saveChanges:=False
    End If
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickImpMulti: "; errNumber; errDesc
        MsgBox "Unable to import table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    ElseIf Not hasIdColumn Then
        MsgBox MissingIdMessage(), vbInformation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Export the T_Multi table to a new workbook in a user-selected folder.")
'@EntryPoint
Public Sub clickExportMulti(ByRef control As IRibbonControl)
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim lo As ListObject
    Dim table As CustomTable
    Dim exportBook As Workbook
    Dim exportSheet As Worksheet
    Dim folderPath As String
    Dim exportPath As String

    'Show folder picker before entering busy state
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        GoTo Cleanup
    End If

    Set table = CustomTable.Create(lo)

    'Create a new workbook and export the table
    Set exportBook = Workbooks.Add
    Set exportSheet = exportBook.Worksheets(1)
    table.Export sh:=exportSheet, startLine:=1, startColumn:=1, addListObject:=True

    'Build the export file path with timestamp
    folderPath = io.Folder()
    If Right$(folderPath, 1) <> Application.PathSeparator Then
        folderPath = folderPath & Application.PathSeparator
    End If
    exportPath = folderPath & TABLE_MULTI & "_export_" & _
                 Format$(Now, "yyyymmdd\_hhnnss") & ".xlsx"

    exportBook.SaveAs Filename:=exportPath, FileFormat:=xlOpenXMLWorkbook
    exportBook.Close saveChanges:=False
    Set exportBook = Nothing

    appScope.Restore
    Set appScope = Nothing
    MsgBox "Exported to: " & exportPath, vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not exportBook Is Nothing Then
        exportBook.Close saveChanges:=False
    End If
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickExportMulti: "; errNumber; errDesc
        MsgBox "Unable to export table: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Generate one linelist per filled row of the T_Multi table.")
'@details
'The multi driver. One generation report serves the whole run and one
'progress bar moves over the rows when the sheet carries its range. The
'summary message closes the run; each row's own outcome is in the
'result column.
'@EntryPoint
Public Sub clickGenerateMulti(ByRef control As IRibbonControl)
    Dim lo As ListObject
    Dim appScope As ApplicationState
    Dim entry As DesignerEntry
    Dim bar As ProgressBar
    Dim buildRows As Long
    Dim builtCount As Long
    Dim failedCount As Long

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    buildRows = CountBuildRows(lo)
    If buildRows = 0 Then
        MsgBox "No row carries a setup file. Fill the " & Chr(34) & COL_SETUPS & _
               Chr(34) & " column first.", vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlNorthWestArrow

    Set entry = EventsDesignerCore.EntryManager()

    'One report for the whole run; every row flushes into it
    GenerationReport.InitReport ThisWorkbook

    Set bar = ResolveProgressBar(lo.Parent, buildRows)

    GenerateMultipleRows lo, entry, bar, builtCount, failedCount

    If Not bar Is Nothing Then bar.Complete CStr(builtCount) & " built"
    GenerationReport.FinaliseReport

    appScope.Restore
    Set appScope = Nothing
    MsgBox CStr(builtCount) & " linelist(s) built, " & CStr(failedCount) & _
           " failed. The " & Chr(34) & COL_RESULT & Chr(34) & _
           " column carries each row's outcome.", _
           vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'Show whatever report was written before the error
    GenerationReport.FinaliseReport
    If Not appScope Is Nothing Then appScope.Restore
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerateMulti: "; errNumber; errDesc
        MsgBox "Unable to run the multi generation: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Multi generation driver
'===============================================================================

'@Description("Walk the T_Multi rows and run one build per filled row.")
'@details
'A row builds when its setups cell is filled. The row's values land on
'Main through the entry, the entry checks run, and the build follows;
'every bundle flushes into the one report of the run. A row that is
'refused or fails writes its fault into the result column and the loop
'keeps running. A row with content and an empty setups cell reports
'itself skipped in the result column; a fully empty row stays untouched.
'The table and the entry arrive as parameters so a suite can drive the
'loop with fixture objects.
'@param lo ListObject. The T_Multi ListObject.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param bar ProgressBar. The bar over the rows. Nothing means no bar.
'@param builtCount Long. Answers the number of rows built and saved.
'@param failedCount Long. Answers the number of rows refused or failed.
Public Sub GenerateMultipleRows(ByVal lo As ListObject, _
                                ByVal entry As DesignerEntry, _
                                ByVal bar As ProgressBar, _
                                ByRef builtCount As Long, _
                                ByRef failedCount As Long)
    Dim rowIdx As Long
    Dim processed As Long
    Dim setupPath As String
    Dim outcomeText As String

    builtCount = 0
    failedCount = 0

    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIdx = 1 To lo.ListRows.Count
        setupPath = CellText(lo, rowIdx, COL_SETUPS)

        If LenB(setupPath) > 0 Then
            processed = processed + 1
            If Not bar Is Nothing Then bar.Update processed - 1, BaseName(setupPath)

            FlushRowHeader lo, rowIdx, setupPath
            WriteRowEntries lo, rowIdx, entry

            If BuildRow(entry, outcomeText) Then
                builtCount = builtCount + 1
                WriteRowCell lo, rowIdx, COL_OUTPUT_FILES, outcomeText
                WriteRowCell lo, rowIdx, COL_RESULT, RESULT_BUILT
            Else
                failedCount = failedCount + 1
                WriteRowCell lo, rowIdx, COL_RESULT, outcomeText
            End If

            If Not bar Is Nothing Then bar.Update processed, BaseName(setupPath)
        ElseIf RowHasContent(lo, rowIdx) Then
            WriteRowCell lo, rowIdx, COL_RESULT, _
                         "Skipped: the " & COL_SETUPS & " cell is empty."
        End If
    Next rowIdx
End Sub

'@Description("Write one row's values into the Main entries through the entry manager.")
'@details
'Every mapped column is written, blanks included, so a row starts from
'its own values alone and behaves like a Main the user typed: an empty
'geobase is the optional entry, an empty epiweek reads as week 1, an
'empty required value is refused by the entry checks with the row's own
'message. The output files cell may carry the full path the last run
'wrote; OutputNameFromCell reduces it to the file name, which keeps a
're-run of the table stable.
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
Public Sub WriteRowEntries(ByVal lo As ListObject, _
                           ByVal rowIdx As Long, _
                           ByVal entry As DesignerEntry)
    entry.AddInfo CellText(lo, rowIdx, COL_SETUPS), "setuppath"
    entry.AddInfo CellText(lo, rowIdx, COL_GEOBASES), "geopath"
    entry.AddInfo CellText(lo, rowIdx, COL_OUTPUT_FOLDERS), "lldir"
    entry.AddInfo OutputNameFromCell(CellText(lo, rowIdx, COL_OUTPUT_FILES)), "llname"
    entry.AddInfo CellText(lo, rowIdx, COL_PASSWORD), "llpassword"
    entry.AddInfo CellText(lo, rowIdx, COL_DEBUG_PASSWORD), "debugpassword"
    entry.AddInfo CellText(lo, rowIdx, COL_LANG_DICTIONARY), "setuplang"
    entry.AddInfo CellText(lo, rowIdx, COL_LANG_INTERFACE), "lllang"
    entry.AddInfo CellText(lo, rowIdx, COL_EPIWEEK_START), "epiweekstart"
    entry.AddInfo CellText(lo, rowIdx, COL_DESIGN), "design"
End Sub

'@Description("Reduce an output files cell to the linelist file name.")
'@details
'A row that built gets the full written path in its output files cell,
'and a re-run reads that cell back. The folder part and the .xlsb
'extension go, so the name reaches the build the same on the first run
'and on every re-run. Both path separators are handled; a table filled
'on one platform stays readable on the other.
'@param cellValue String. The output files cell text.
'@return String. The bare file name.
Public Function OutputNameFromCell(ByVal cellValue As String) As String
    Dim nameText As String

    nameText = BaseName(Trim$(cellValue))

    If LCase$(Right$(nameText, 5)) = ".xlsb" Then
        nameText = Left$(nameText, Len(nameText) - 5)
    End If

    OutputNameFromCell = nameText
End Function

'@Description("Count the rows whose setups cell is filled.")
'@param lo ListObject. The T_Multi ListObject.
'@return Long. The number of rows the driver will build.
Public Function CountBuildRows(ByVal lo As ListObject) As Long
    Dim rowIdx As Long

    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIdx = 1 To lo.ListRows.Count
        If LenB(CellText(lo, rowIdx, COL_SETUPS)) > 0 Then
            CountBuildRows = CountBuildRows + 1
        End If
    Next rowIdx
End Function

'@Description("Run the entry checks and one build; answer the outcome with no raise.")
'@details
'The one place a row's fault is caught, so the loop keeps running. A
'refused row answers False with a pointer at the report, where the
'checks landed; a failed row answers False with the error text, and the
'incomplete output workbook is discarded quietly.
'@param entry DesignerEntry. The entry manager, already loaded with the row.
'@param outcomeText String. Answers the written path on success and the fault text otherwise.
'@return Boolean. True when the row built and saved.
Private Function BuildRow(ByVal entry As DesignerEntry, ByRef outcomeText As String) As Boolean
    Dim ll As Linelist

    On Error GoTo Fail

    If Not EventsDesignerAdvanced.ValidateEntries(entry) Then
        outcomeText = "Refused by the entry checks. Open the generation report."
        Exit Function
    End If

    outcomeText = EventsDesignerAdvanced.GenerateOne(entry, ll)
    BuildRow = True
    Exit Function

Fail:
    outcomeText = "Failed: " & Err.Description

    'The failed build leaves its output workbook open; the batch closes
    'it quietly and moves to the next row.
    On Error Resume Next
    If Not ll Is Nothing Then ll.DiscardBuild
    On Error GoTo 0
End Function

'@Description("Flush one report bundle naming the row before its build starts.")
'@details
'The bundles of every row land in the one report of the run, and this
'line tells the reader which row the entries under it belong to.
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param setupPath String. The row's setup file path.
Private Sub FlushRowHeader(ByVal lo As ListObject, _
                           ByVal rowIdx As Long, _
                           ByVal setupPath As String)
    Dim rowChecks As Checking
    Dim batch As BetterArray
    Dim rowId As String

    rowId = CellText(lo, rowIdx, COL_ID)
    If LenB(rowId) = 0 Then rowId = "row " & CStr(rowIdx)

    Set rowChecks = Checking.Create("Multi build " & rowId, setupPath)
    rowChecks.Add rowId, "Build started for: " & setupPath, checkingInfo

    Set batch = New BetterArray
    batch.LowerBound = 1
    batch.Push rowChecks
    GenerationReport.FlushCheckings batch
End Sub

'@Description("Build the bar over the rows when the sheet carries the named ranges.")
'@details
'The bar range and the status cell are owner hand work on the mock. A
'missing name means no bar and no raise; the generation runs the same.
'A name that resolves outside the GenerateMultiple sheet is treated as
'missing, so a later bar on Main keeps its own range.
'@param multiSheet Worksheet. The GenerateMultiple worksheet.
'@param maximum Long. The number of rows the run will build.
'@return ProgressBar. The bar, or Nothing when the range is missing.
Private Function ResolveProgressBar(ByVal multiSheet As Worksheet, _
                                    ByVal maximum As Long) As ProgressBar
    Dim barRange As Range
    Dim statusRange As Range
    Dim bar As ProgressBar

    On Error Resume Next
    Set barRange = multiSheet.Range(RNG_PROGRESS_BAR)
    Set statusRange = multiSheet.Range(RNG_PROGRESS_STATUS)
    On Error GoTo 0

    If barRange Is Nothing Then Exit Function
    If Not barRange.Worksheet Is multiSheet Then Exit Function

    'A malformed hand-made range stops the bar alone; the generation runs on
    On Error Resume Next
    Set bar = ProgressBar.Create(barRange, maximum)
    If Not statusRange Is Nothing Then bar.AttachStatusCell statusRange
    On Error GoTo 0

    Set ResolveProgressBar = bar
End Function

'@Description("Read one cell of a row by column header. A missing column reads as empty.")
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param colName String. The column header.
'@return String. The trimmed cell text.
Private Function CellText(ByVal lo As ListObject, _
                          ByVal rowIdx As Long, _
                          ByVal colName As String) As String
    Dim col As ListColumn

    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then Exit Function
    CellText = Trim$(CStr(lo.ListRows(rowIdx).Range.Cells(1, col.Index).Value))
End Function

'@Description("Write one cell of a row by column header. A missing column skips the write.")
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param colName String. The column header.
'@param cellValue String. The value to write.
Private Sub WriteRowCell(ByVal lo As ListObject, _
                         ByVal rowIdx As Long, _
                         ByVal colName As String, _
                         ByVal cellValue As String)
    Dim col As ListColumn

    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then Exit Sub
    lo.ListRows(rowIdx).Range.Cells(1, col.Index).Value = cellValue
End Sub

'@Description("True when any mapped cell of the row apart from setups is filled.")
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@return Boolean. True when the row carries content.
Private Function RowHasContent(ByVal lo As ListObject, ByVal rowIdx As Long) As Boolean
    Dim mappedCols As Variant
    Dim idx As Long

    mappedCols = Array(COL_GEOBASES, COL_OUTPUT_FOLDERS, COL_OUTPUT_FILES, _
                       COL_PASSWORD, COL_DEBUG_PASSWORD, COL_LANG_DICTIONARY, _
                       COL_LANG_INTERFACE, COL_EPIWEEK_START, COL_DESIGN)

    For idx = LBound(mappedCols) To UBound(mappedCols)
        If LenB(CellText(lo, rowIdx, CStr(mappedCols(idx)))) > 0 Then
            RowHasContent = True
            Exit Function
        End If
    Next idx
End Function

'@Description("The file name part of a path.")
'@param filePath String. A file path or a bare name.
'@return String. The text after the last path separator.
Private Function BaseName(ByVal filePath As String) As String
    Dim sepPos As Long
    Dim altPos As Long

    sepPos = InStrRev(filePath, "/")
    altPos = InStrRev(filePath, "\")
    If altPos > sepPos Then sepPos = altPos

    BaseName = Mid$(filePath, sepPos + 1)
End Function


'@section Table and dropdown resolution
'===============================================================================

'@Description("Resolve the T_Multi ListObject from the GenerateMultiple worksheet.")
'@param targetBook Optional Workbook. The workbook to resolve on. Defaults to this workbook.
'@return ListObject. The T_Multi ListObject, or Nothing when not found.
Public Function ResolveMultiTable(Optional ByVal targetBook As Workbook = Nothing) As ListObject
    Dim sh As Worksheet

    If targetBook Is Nothing Then Set targetBook = ThisWorkbook

    On Error Resume Next
    Set sh = targetBook.Worksheets(SHEET_GENERATE_MULTIPLE)
    On Error GoTo 0

    If sh Is Nothing Then Exit Function

    On Error Resume Next
    Set ResolveMultiTable = sh.ListObjects(TABLE_MULTI)
    On Error GoTo 0
End Function

'@Description("Resolve the DropdownLists manager on the __dropdowns worksheet.")
'@param targetBook Optional Workbook. The workbook to resolve on. Defaults to this workbook.
'@return DropdownLists. The dropdown manager, or Nothing when the sheet is missing.
Public Function ResolveDropdownManager(Optional ByVal targetBook As Workbook = Nothing) As DropdownLists
    Dim dropSheet As Worksheet

    If targetBook Is Nothing Then Set targetBook = ThisWorkbook

    On Error Resume Next
    Set dropSheet = targetBook.Worksheets(SHEET_DROPDOWNS)
    On Error GoTo 0

    If dropSheet Is Nothing Then Exit Function

    Set ResolveDropdownManager = DropdownLists.Create(dropSheet, DROPDOWN_PREFIX)
End Function

'@Description("Fill only the blank ID cells with the next free numbers.")
'@details
'An ID is written once. This scans the ID column for the largest number
'already written, then gives every blank cell the next numbers, so a row
'keeps the dropdown named after its ID for its whole life.
'@param lo ListObject. The T_Multi ListObject.
'@return Boolean. True when the ID column exists.
Public Function EnsureRowIds(ByVal lo As ListObject) As Boolean
    Dim idCol As ListColumn
    Dim idRange As Range
    Dim rowIdx As Long
    Dim nextNumber As Long
    Dim numberPart As Long
    Dim cellText As String

    On Error Resume Next
    Set idCol = lo.ListColumns(COL_ID)
    On Error GoTo 0

    If idCol Is Nothing Then Exit Function
    EnsureRowIds = True

    Set idRange = idCol.DataBodyRange
    If idRange Is Nothing Then Exit Function

    'Find the largest number already written
    For rowIdx = 1 To idRange.Rows.Count
        cellText = Trim$(CStr(idRange.Cells(rowIdx, 1).Value))
        If LenB(cellText) > 0 Then
            numberPart = TrailingNumber(cellText)
            If numberPart > nextNumber Then nextNumber = numberPart
        End If
    Next rowIdx

    'Fill the blank cells. The ID shape matches the one CustomTable
    'writes: the prefix, one space, the number.
    For rowIdx = 1 To idRange.Rows.Count
        cellText = Trim$(CStr(idRange.Cells(rowIdx, 1).Value))
        If LenB(cellText) = 0 Then
            nextNumber = nextNumber + 1
            idRange.Cells(rowIdx, 1).Value = ID_PREFIX & " " & CStr(nextNumber)
        End If
    Next rowIdx
End Function

'@Description("Read the number at the end of an ID value.")
'@param idText String. The ID cell text.
'@return Long. The trailing number, or 0 when the text ends without digits.
Private Function TrailingNumber(ByVal idText As String) As Long
    Dim charIdx As Long
    Dim oneChar As String
    Dim digits As String

    For charIdx = Len(idText) To 1 Step -1
        oneChar = Mid$(idText, charIdx, 1)
        If oneChar Like "[0-9]" Then
            digits = oneChar & digits
        Else
            Exit For
        End If
    Next charIdx

    If LenB(digits) > 0 Then TrailingNumber = CLng(digits)
End Function

'@Description("Return the T_Multi column header matching the active cell position.")
'@param lo ListObject. The T_Multi ListObject.
'@return String. Column header name, or vbNullString when outside the table.
Private Function ActiveCellColumnName(ByVal lo As ListObject) As String
    Dim colOffset As Long

    If Intersect(Application.ActiveCell, lo.Range) Is Nothing Then Exit Function

    colOffset = Application.ActiveCell.Column - lo.HeaderRowRange.Column + 1
    If colOffset < 1 Or colOffset > lo.ListColumns.Count Then Exit Function

    ActiveCellColumnName = lo.ListColumns(colOffset).Name
End Function


'@section Folder multi helpers -- file loading by column type
'===============================================================================

'@Description("Write setup paths into the setups column and wire one language dropdown per row.")
'@param lo ListObject. The T_Multi ListObject.
'@param filePaths BetterArray. Setup file paths (1-based).
'@param startRow Long. Worksheet row number the first path lands on.
'@param drop DropdownLists. The dropdown manager of the host workbook.
'@param skipped BetterArray. Collects one line per skipped step.
Public Sub LoadSetupFiles(ByVal lo As ListObject, _
                          ByVal filePaths As BetterArray, _
                          ByVal startRow As Long, _
                          ByVal drop As DropdownLists, _
                          ByVal skipped As BetterArray)
    Dim setupBook As Workbook
    Dim tradSheet As Worksheet
    Dim langValues As BetterArray
    Dim langCol As ListColumn
    Dim idCol As ListColumn
    Dim currentRow As Long
    Dim filePath As String
    Dim dropName As String
    Dim rowId As String
    Dim langCell As Range
    Dim idx As Long

    If filePaths Is Nothing Then Exit Sub
    If filePaths.Length = 0 Then Exit Sub

    WriteFilesToColumn lo, COL_SETUPS, startRow, filePaths, skipped

    'The write may have extended the table; the new rows take their IDs
    'here so each row's dropdown has its name.
    If Not EnsureRowIds(lo) Then
        skipped.Push MissingIdMessage()
        Exit Sub
    End If

    On Error Resume Next
    Set langCol = lo.ListColumns(COL_LANG_DICTIONARY)
    Set idCol = lo.ListColumns(COL_ID)
    On Error GoTo 0

    If langCol Is Nothing Then
        skipped.Push MissingColumnMessage(COL_LANG_DICTIONARY) & _
                     " The language dropdowns were skipped."
        Exit Sub
    End If

    If drop Is Nothing Then
        skipped.Push "The " & SHEET_DROPDOWNS & _
                     " sheet is missing. The language dropdowns were skipped."
        Exit Sub
    End If

    'For each setup file, extract languages and wire a per-row dropdown
    currentRow = startRow
    For idx = filePaths.LowerBound To filePaths.UpperBound
        filePath = CStr(filePaths.Item(idx))

        'Open the setup file read-only
        Set setupBook = Nothing
        On Error Resume Next
        Set setupBook = Workbooks.Open(filePath, ReadOnly:=True)
        On Error GoTo 0

        If setupBook Is Nothing Then
            skipped.Push "This setup file failed to open: " & filePath
        Else
            'Resolve the Translations worksheet
            Set tradSheet = Nothing
            On Error Resume Next
            Set tradSheet = setupBook.Worksheets(SHEET_TRANSLATIONS)
            On Error GoTo 0

            If tradSheet Is Nothing Then
                skipped.Push "This setup file has no " & SHEET_TRANSLATIONS & _
                             " sheet: " & filePath
            Else
                Set langValues = EventsDesignerAdvanced.SetupLanguages(tradSheet)
                If langValues.Length = 0 Then
                    skipped.Push "No language was found in: " & filePath
                Else
                    'The dropdown is named after the row ID
                    rowId = CStr(lo.Parent.Cells(currentRow, idCol.Range.Column).Value)
                    dropName = rowId & LANG_SUFFIX

                    'Add or update the dropdown with extracted languages
                    If drop.Exists(dropName) Then
                        drop.Update langValues, dropName
                    Else
                        drop.Add langValues, dropName
                    End If

                    'Apply validation on the language cell using the dropdown
                    Set langCell = lo.Parent.Cells(currentRow, langCol.Range.Column)
                    drop.SetValidation langCell, dropName
                End If
            End If

            setupBook.Close saveChanges:=False
            Set setupBook = Nothing
        End If

        currentRow = currentRow + 1
    Next idx
End Sub

'@Description("Write geobase paths into the geobases column.")
'@param lo ListObject. The T_Multi ListObject.
'@param filePaths BetterArray. Geobase file paths (1-based).
'@param startRow Long. Worksheet row number the first path lands on.
'@param skipped BetterArray. Collects one line per skipped step.
Public Sub LoadGeobaseFiles(ByVal lo As ListObject, _
                            ByVal filePaths As BetterArray, _
                            ByVal startRow As Long, _
                            ByVal skipped As BetterArray)
    If filePaths Is Nothing Then Exit Sub
    If filePaths.Length = 0 Then Exit Sub

    WriteFilesToColumn lo, COL_GEOBASES, startRow, filePaths, skipped
    If Not EnsureRowIds(lo) Then skipped.Push MissingIdMessage()
End Sub

'@Description("Write a folder path into the output folders column at the given row.")
'@param lo ListObject. The T_Multi ListObject.
'@param folderPath String. The selected folder path.
'@param startRow Long. Worksheet row number to write on.
'@param skipped BetterArray. Collects one line per skipped step.
Public Sub LoadOutputFolder(ByVal lo As ListObject, _
                            ByVal folderPath As String, _
                            ByVal startRow As Long, _
                            ByVal skipped As BetterArray)
    Dim col As ListColumn

    On Error Resume Next
    Set col = lo.ListColumns(COL_OUTPUT_FOLDERS)
    On Error GoTo 0

    If col Is Nothing Then
        skipped.Push MissingColumnMessage(COL_OUTPUT_FOLDERS)
        Exit Sub
    End If

    lo.Parent.Cells(startRow, col.Range.Column).Value = folderPath
End Sub

'@Description("Write file paths into a column, adding rows to the table as needed.")
'@param lo ListObject. The T_Multi ListObject.
'@param colName String. Column header to write into.
'@param startRow Long. Worksheet row number to start writing from.
'@param filePaths BetterArray. File paths to write (1-based).
'@param skipped BetterArray. Collects one line per skipped step.
Private Sub WriteFilesToColumn(ByVal lo As ListObject, _
                               ByVal colName As String, _
                               ByVal startRow As Long, _
                               ByVal filePaths As BetterArray, _
                               ByVal skipped As BetterArray)
    Dim col As ListColumn
    Dim currentRow As Long
    Dim lastDataRow As Long
    Dim idx As Long

    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then
        skipped.Push MissingColumnMessage(colName)
        Exit Sub
    End If

    currentRow = startRow

    For idx = filePaths.LowerBound To filePaths.UpperBound
        'Ensure the row exists within the table
        lastDataRow = lo.HeaderRowRange.Row + lo.ListRows.Count
        If currentRow > lastDataRow Then
            lo.ListRows.Add
        End If

        lo.Parent.Cells(currentRow, col.Range.Column).Value = CStr(filePaths.Item(idx))
        currentRow = currentRow + 1
    Next idx
End Sub


'@section Message helpers
'===============================================================================

'@Description("Collect the selected file paths into a BetterArray (1-based).")
'@param io OSFiles. The file picker with selected files.
'@return BetterArray. The selected file paths.
Private Function CollectFiles(ByVal io As OSFiles) As BetterArray
    Dim filePaths As BetterArray

    Set filePaths = New BetterArray
    filePaths.LowerBound = 1

    io.ResetFilesIterator
    Do While io.HasNextFile()
        filePaths.Push io.NextFile()
    Loop

    Set CollectFiles = filePaths
End Function

'@Description("Show the collected skip lines in one message.")
'@param skipped BetterArray. The skip lines pushed by the helpers.
Private Sub ShowSkipped(ByVal skipped As BetterArray)
    Dim message As String
    Dim idx As Long

    If skipped.Length = 0 Then Exit Sub

    For idx = skipped.LowerBound To skipped.UpperBound
        message = message & CStr(skipped.Item(idx)) & vbNewLine
    Next idx

    MsgBox "Some steps were skipped:" & vbNewLine & message, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("Tell the user the multi table is missing.")
Private Sub ReportMissingTable()
    MsgBox "The " & TABLE_MULTI & " table was not found on the " & _
           SHEET_GENERATE_MULTIPLE & " sheet.", _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("Build the message for a missing column.")
'@param colName String. The missing column header.
'@return String. The message line.
Private Function MissingColumnMessage(ByVal colName As String) As String
    MissingColumnMessage = "The column " & Chr(34) & colName & Chr(34) & _
                           " is missing on " & TABLE_MULTI & "."
End Function

'@Description("Build the message for a missing ID column.")
'@return String. The message line.
Private Function MissingIdMessage() As String
    MissingIdMessage = "The " & TABLE_MULTI & " table has no " & COL_ID & _
                       " column, so the rows got no ID."
End Function
