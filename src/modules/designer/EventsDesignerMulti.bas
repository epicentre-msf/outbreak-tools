Attribute VB_Name = "EventsDesignerMulti"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the Multi group on the designer workbook, and the driver that generates one linelist per row.")
'@depends CustomTable, ApplicationState, OSFiles, DropdownLists, BetterArray, EventsDesignerAdvanced, EventsDesignerCore, DesignerEntry, DesignerPreparation, GenerationLog, Checking, ProgressBar, GenerationHost, TemporaryRepos
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
'dropdown the validation points at. SafeDropdownName is what turns the
'ID into that name: an ID reads "Operation- 1" and Excel refuses a name
'carrying a dash.
'
'EVERY BUTTON OF THE GROUP WORKS ON ONE WORKSHEET
'-------------------------------------------------------------------------------
'A press answers nothing at all unless the GenerateMultiple sheet of this
'workbook is the one in front. OnMultiSheet is the guard, and it opens
'every callback here.
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
'The screen stays off from the press to the restore, and the report is
'shown once, at the end.
'
'THE RIBBON FILE OF THE RUN
'-------------------------------------------------------------------------------
'The rows used to build with whatever template file was loaded on the
'Main worksheet, which is a screen the user of this table never opens.
'The run now carries its own: the path is kept in the designer's hidden
'names (DesignerPreparation.RibbonTemplatePath) and every row is built
'with it. The first Generate asks for the file, and so does a Generate
'that finds the stored file gone from the disk. clickRibbonMulti is the
'button that changes it or clears it at any time.
'
'THE REPORT OF THE RUN
'-------------------------------------------------------------------------------
'Every row opens one section of the report and closes it. The parts of
'that row's build -- the dictionary, the choices, each data entry sheet,
'the analyses -- become subsections under the row's heading, so a run of
'ten linelists reads as ten blocks, one block per linelist.

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

'The three dropdowns the designer already registers for its own Main
'entries, in DesignerPreparation. The dictionary language is read off each
'row's setup file and so has a list per row; these three are one list for
'the whole workbook, and the rows take them as they are.
Private Const DROP_INTERFACE_LANGUAGES As String = "__interface_languages"
Private Const DROP_EPIWEEK_START As String = "__epiweek_start"
Private Const DROP_DESIGN_VALUES As String = "__design_values"

Private Const MSG_PLACE_DATA As String = "Please place the cursor inside the table data area."

'Progress bar over the rows. The two names live on the GenerateMultiple
'sheet and are owner hand work on the mock; a missing name means no bar
'and no raise, so this code lands before the hand work and wakes up
'with it.
Private Const RNG_PROGRESS_BAR As String = "RNG_ProgressBar"

'The scratch folder TemporaryRepos makes under an output folder, and the
'name of the unfinished workbook a failed build keeps there. Both are
'repeated from the classes that own them (their constants are Private),
'so the pre-flight of a row can name the kept file before any build runs.
Private Const SCRATCH_FOLDER As String = "OBTApp_"
Private Const KEPT_FILE_NAME As String = "__temp.xlsb"
Private Const RNG_PROGRESS_STATUS As String = "RNG_ProgressStatus"

'The result column value of a row that built and saved
Private Const RESULT_BUILT As String = "OK"

'The ribbon file of the run. The filter is the shape of a template
'workbook, the same one the Main worksheet's template button uses.
Private Const TEMPLATE_FILTER As String = "*.xlsb"
Private Const MSG_RIBBON_ASK As String = _
    "Pick the ribbon file the generated linelists are built with."
Private Const MSG_RIBBON_GONE As String = _
    "The ribbon file of this designer was not found on the disk. Pick it again:"
Private Const MSG_RIBBON_NONE As String = _
    "No ribbon file was picked. The generated linelists will carry buttons." & _
    vbNewLine & vbNewLine & "Build them that way?"
Private Const MSG_RIBBON_CURRENT As String = "The ribbon file of this designer is:"
Private Const MSG_RIBBON_CHANGE As String = _
    "Yes picks another file, No clears it and builds with buttons, " & _
    "Cancel keeps the file above."
Private Const MSG_RIBBON_CLEARED As String = _
    "The ribbon file was cleared. The generated linelists will carry buttons."
Private Const MSG_RIBBON_SET As String = "The ribbon file is now:"

'The two bundle titles a row files itself. Under an open section they
'read as the first and the last subsection of the row.
Private Const TITLE_ROW_START As String = "build start"
Private Const TITLE_ROW_OUTCOME As String = "build outcome"


'@section Multi group callbacks
'===============================================================================

'@Description("Load files or folder into the active T_Multi column (setups, geobases, output folders).")
'@EntryPoint
Public Sub clickFolderMulti(ByRef ribbonControl As IRibbonControl)
    Dim lo As ListObject
    Dim colName As String
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim skipped As BetterArray
    Dim startRow As Long

    If Not OnMultiSheet() Then Exit Sub

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

    'The rows the load just added take the shared lists too
    WireSharedColumns lo, ResolveDropdownManager()

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
Public Sub clickDupMulti(ByRef ribbonControl As IRibbonControl)
    Dim lo As ListObject
    Dim appScope As ApplicationState
    Dim relPos As Long
    Dim sourceRow As Range
    Dim destRow As Range
    Dim idCol As ListColumn
    Dim idMissing As Boolean

    If Not OnMultiSheet() Then Exit Sub

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

    WireSharedColumns lo, ResolveDropdownManager()

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
Public Sub clickAddRowsMulti(ByRef ribbonControl As IRibbonControl)
    Dim lo As ListObject
    Dim table As CustomTable
    Dim appScope As ApplicationState
    Dim hasIdColumn As Boolean

    If Not OnMultiSheet() Then Exit Sub

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

    WireSharedColumns lo, ResolveDropdownManager()

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
Public Sub clickResizeMulti(ByRef ribbonControl As IRibbonControl)
    Dim lo As ListObject
    Dim table As CustomTable
    Dim appScope As ApplicationState
    Dim hasIdColumn As Boolean

    If Not OnMultiSheet() Then Exit Sub

    hasIdColumn = True

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    'THE THRESHOLD COUNTS THE ID. RemoveRows drops a row whose filled cell
    'count is at or below the number it is given, and 0 asks it to work that
    'number out from the formula columns. T_Multi carries none, so the answer
    'was 0 and a row had to hold nothing at all to go. Every row carries an
    'ID, written once and never rewritten, so no row was ever that empty and
    'the button did nothing.
    Set table = CustomTable.Create(lo)
    table.RemoveRows totalCount:=1, includeIds:=False, forceShift:=False
    hasIdColumn = EnsureRowIds(lo)

    WireSharedColumns lo, ResolveDropdownManager()

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
Public Sub clickImpMulti(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim importBook As Workbook
    Dim sourceLo As ListObject
    Dim targetLo As ListObject
    Dim sourceTable As CustomTable
    Dim targetTable As CustomTable
    Dim hasIdColumn As Boolean

    If Not OnMultiSheet() Then Exit Sub

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

    WireSharedColumns targetLo, ResolveDropdownManager()

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
Public Sub clickExportMulti(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles
    Dim appScope As ApplicationState
    Dim lo As ListObject
    Dim table As CustomTable
    Dim exportBook As Workbook
    Dim exportSheet As Worksheet
    Dim folderPath As String
    Dim exportPath As String

    If Not OnMultiSheet() Then Exit Sub

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

'@Description("Show the ribbon file of the designer and change or clear it.")
'@details
'The one button that manages the ribbon file the multi generation builds
'with. A designer that already holds one shows it and offers the three
'answers: pick another file, clear it so the linelists carry buttons, or
'keep what is there. A designer that holds none goes straight to the
'picker.
'
'Generate asks for the file on its own when the designer holds none, so
'this button is for the day the file moves or the choice changes.
'@EntryPoint
Public Sub clickRibbonMulti(ByRef ribbonControl As IRibbonControl)
    Dim prep As DesignerPreparation
    Dim currentPath As String
    Dim answer As VbMsgBoxResult
    Dim pickedPath As String

    If Not OnMultiSheet() Then Exit Sub

    On Error GoTo Cleanup
    Set prep = DesignerPreparation.Create(ThisWorkbook)
    currentPath = prep.RibbonTemplatePath

    If LenB(currentPath) > 0 Then
        answer = MsgBox(MSG_RIBBON_CURRENT & vbNewLine & currentPath & vbNewLine & _
                        vbNewLine & MSG_RIBBON_CHANGE, _
                        vbQuestion + vbYesNoCancel, PROMPT_TITLE)

        If answer = vbCancel Then Exit Sub

        If answer = vbNo Then
            prep.RibbonTemplatePath = vbNullString
            MsgBox MSG_RIBBON_CLEARED, vbInformation + vbOKOnly, PROMPT_TITLE
            Exit Sub
        End If
    End If

    pickedPath = AskRibbonTemplate(prep)
    If LenB(pickedPath) = 0 Then Exit Sub

    MsgBox MSG_RIBBON_SET & vbNewLine & pickedPath, _
           vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Debug.Print "clickRibbonMulti: "; Err.Number; Err.Description
    MsgBox "Unable to read the ribbon file: " & Err.Description, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("Generate one linelist per filled row of the T_Multi table.")
'@details
'The multi driver. One generation report serves the whole run and one
'progress bar moves over the rows when the sheet carries its range. The
'report is shown once the run is over, then the summary message closes
'the run; each row's own outcome is in the result column.
'
'The ribbon file of the run is settled before anything else, because it
'may open a dialog. Every row is built with it. The path is read next,
'through GenerationHost.InPlace: in place the rows build in this Excel
'with the screen off for the whole run; on the instance path they build
'in one hidden Excel over one copy of the designer, and the bar over the
'rows moves here with the screen on.
'@EntryPoint
Public Sub clickGenerateMulti(ByRef ribbonControl As IRibbonControl)
    Dim lo As ListObject
    Dim buildRows As Long
    Dim templatePath As String
    Dim designerBook As Workbook
    Dim host As GenerationHost

    If Not OnMultiSheet() Then Exit Sub

    Set lo = ResolveMultiTable()
    If lo Is Nothing Then
        ReportMissingTable
        Exit Sub
    End If

    Set designerBook = lo.Parent.Parent

    buildRows = CountBuildRows(lo)
    If buildRows = 0 Then
        MsgBox "No row carries a setup file. Fill the " & Chr(34) & COL_SETUPS & _
               Chr(34) & " column first.", vbInformation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    'The ribbon file every row of this run is built with. A designer that
    'holds one keeps it; a designer that holds none, or one whose file has
    'left the disk, asks here. A user who wants none says so and the run
    'goes on with the buttons.
    If Not ResolveRibbonTemplate(designerBook, templatePath) Then Exit Sub

    On Error GoTo Cleanup
    Set host = GenerationHost.Create(designerBook)

    If host.InPlace Then
        Set host = Nothing
        GenerateMultiInPlace lo, buildRows, templatePath
    Else
        GenerateMultiInInstance lo, designerBook, buildRows, templatePath, host
    End If
    Exit Sub

Cleanup:
    Debug.Print "clickGenerateMulti: "; Err.Number; Err.Description
    MsgBox "Unable to run the multi generation: " & Err.Description, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
End Sub

'@Description("The multi run in this Excel, with the screen off from the press to the restore.")
'@details
'The A3 driver: one busy scope over the whole run, the rows through
'GenerateMultipleRows, the restore, the report once.
'@param lo ListObject. The T_Multi ListObject.
'@param buildRows Long. The number of rows the run builds, for the bar.
'@param templatePath String. The ribbon file of the run. Empty builds the buttons.
Private Sub GenerateMultiInPlace(ByVal lo As ListObject, _
                                 ByVal buildRows As Long, _
                                 ByVal templatePath As String)
    Dim appScope As ApplicationState
    Dim entry As DesignerEntry
    Dim bar As ProgressBar
    Dim builtCount As Long
    Dim failedCount As Long

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlNorthWestArrow

    Set entry = EventsDesignerCore.EntryManager()

    'One run log for the whole run; every row flushes into it. The log
    'opens bare: each row names itself in its own header bundle.
    EventsDesignerAdvanced.StartRunLog

    Set bar = ResolveProgressBar(lo.Parent, buildRows)

    GenerateMultipleRows lo, entry, bar, builtCount, failedCount, templatePath

    If Not bar Is Nothing Then bar.Complete CStr(builtCount) & " built"
    EventsDesignerAdvanced.FinishRunLog RunSummary(builtCount, failedCount)

    appScope.Restore
    Set appScope = Nothing

    'The report, once, with the screen already back on.
    EventsDesignerAdvanced.ShowRunLog
    MsgBox RunSummary(builtCount, failedCount) & ". The " & Chr(34) & COL_RESULT & _
           Chr(34) & " column carries each row's outcome.", _
           vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'Close the log over whatever was logged before the error, and show it
    EventsDesignerAdvanced.FinishRunLog "Failed: " & errDesc
    If Not appScope Is Nothing Then appScope.Restore
    EventsDesignerAdvanced.ShowRunLog
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerateMulti: "; errNumber; errDesc
        MsgBox "Unable to run the multi generation: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("The multi run in the hidden instance, with the bar over the rows moving here.")
'@details
'One host and one copy of the designer serve the whole run; every row
'writes its entries on Main here and hands them to the copy as text with
'its first step. This Excel is never busy: events go off and the cursor
'changes, the screen stays on, and the bar's cell writes paint on their
'own. The copy is written in the scratch folder beside the designer,
'through a throwaway repository whose drop empties that folder, so the
'kept name is marked on it first. The instance is released on every
'exit, and in the cleanup the host is dropped before anything else runs.
'@param lo ListObject. The T_Multi ListObject.
'@param designerBook Workbook. The designer holding the table.
'@param buildRows Long. The number of rows the run builds, for the bar.
'@param templatePath String. The ribbon file of the run. Empty builds the buttons.
'@param host GenerationHost. The host, created and on the instance path.
Private Sub GenerateMultiInInstance(ByVal lo As ListObject, _
                                    ByVal designerBook As Workbook, _
                                    ByVal buildRows As Long, _
                                    ByVal templatePath As String, _
                                    ByVal host As GenerationHost)
    Dim entry As DesignerEntry
    Dim bar As ProgressBar
    Dim scratch As TemporaryRepos
    Dim builtCount As Long
    Dim failedCount As Long
    Dim previousEvents As Boolean
    Dim previousCursor As Long
    Dim sideHeld As Boolean

    On Error GoTo Cleanup
    Set entry = EventsDesignerCore.EntryManager()

    previousEvents = Application.EnableEvents
    previousCursor = Application.Cursor
    Application.EnableEvents = False
    Application.Cursor = xlWait
    sideHeld = True

    EventsDesignerAdvanced.StartRunLog

    Set bar = ResolveProgressBar(lo.Parent, buildRows)

    Set scratch = TemporaryRepos.Create(designerBook.Path)
    scratch.EnsureReady
    scratch.KeepFile KEPT_FILE_NAME

    host.Acquire
    host.OpenDesignerCopy scratch.RootPath

    GenerateMultipleRows lo, entry, bar, builtCount, failedCount, templatePath, host

    If Not bar Is Nothing Then bar.Complete CStr(builtCount) & " built"
    EventsDesignerAdvanced.FinishRunLog RunSummary(builtCount, failedCount)

    EventsDesignerAdvanced.ReleaseBuildHost host
    Set host = Nothing

    RestoreVisibleSide previousEvents, previousCursor

    EventsDesignerAdvanced.ShowRunLog
    MsgBox RunSummary(builtCount, failedCount) & ". The " & Chr(34) & COL_RESULT & _
           Chr(34) & " column carries each row's outcome.", _
           vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    EventsDesignerAdvanced.ReleaseBuildHost host
    Set host = Nothing
    If Not bar Is Nothing Then bar.Reset errDesc
    EventsDesignerAdvanced.FinishRunLog "Failed: " & errDesc
    If sideHeld Then RestoreVisibleSide previousEvents, previousCursor
    EventsDesignerAdvanced.ShowRunLog
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerateMulti: "; errNumber; errDesc
        MsgBox "Unable to run the multi generation: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("The closing line of a run: how many rows built and how many failed.")
'@param builtCount Long. The rows built and saved.
'@param failedCount Long. The rows refused or failed.
'@return String. "<n> linelist(s) built, <m> failed".
Private Function RunSummary(ByVal builtCount As Long, ByVal failedCount As Long) As String
    RunSummary = CStr(builtCount) & " linelist(s) built, " & CStr(failedCount) & " failed"
End Function

'@Description("Put the events and the cursor of this Excel back.")
'@param previousEvents Boolean. The events setting before the run.
'@param previousCursor Long. The cursor before the run.
Private Sub RestoreVisibleSide(ByVal previousEvents As Boolean, ByVal previousCursor As Long)
    On Error Resume Next
    Application.EnableEvents = previousEvents
    Application.Cursor = previousCursor
    On Error GoTo 0
End Sub


'@section The ribbon file of the run
'===============================================================================

'@Description("Settle the ribbon file of a run, asking for it when the designer holds none.")
'@details
'The stored path wins whenever the file behind it is on the disk, so the
'question is asked once and the answer outlives the session. A stored
'path whose file has gone says so and asks again, which is the other
'case the user wants to hear about.
'
'A picker the user closes is a real answer: an empty template path builds
'the linelists with buttons, and that is offered plainly. Saying no to
'that stops the run before anything is built.
'@param designerBook Workbook. The designer holding the stored path.
'@param templatePath String. Answers the path the run builds with, empty for the buttons.
'@return Boolean. True when the run may start.
Private Function ResolveRibbonTemplate(ByVal designerBook As Workbook, _
                                       ByRef templatePath As String) As Boolean
    Dim prep As DesignerPreparation
    Dim storedPath As String

    templatePath = vbNullString
    Set prep = DesignerPreparation.Create(designerBook)
    storedPath = prep.RibbonTemplatePath

    If LenB(storedPath) > 0 Then
        If LenB(Dir(storedPath)) > 0 Then
            templatePath = storedPath
            ResolveRibbonTemplate = True
            Exit Function
        End If

        MsgBox MSG_RIBBON_GONE & vbNewLine & storedPath, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    Else
        MsgBox MSG_RIBBON_ASK, vbInformation + vbOKOnly, PROMPT_TITLE
    End If

    templatePath = AskRibbonTemplate(prep)

    If LenB(templatePath) > 0 Then
        ResolveRibbonTemplate = True
        Exit Function
    End If

    ResolveRibbonTemplate = _
        (MsgBox(MSG_RIBBON_NONE, vbQuestion + vbYesNo, PROMPT_TITLE) = vbYes)
End Function

'@Description("Show the file picker and store what the user picked.")
'@details
'The one place the ribbon file reaches the hidden names, so the picker
'and the store never fall apart. A closed picker stores nothing and
'answers an empty path.
'@param prep DesignerPreparation. The preparation over the designer workbook.
'@return String. The picked path, empty when the user closed the picker.
Private Function AskRibbonTemplate(ByVal prep As DesignerPreparation) As String
    Dim io As OSFiles

    Set io = OSFiles.Create()
    io.LoadFile TEMPLATE_FILTER

    If Not io.HasValidFile() Then Exit Function

    AskRibbonTemplate = io.File()
    prep.RibbonTemplatePath = AskRibbonTemplate
End Function


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
'
'ONE SECTION PER ROW
'-------------------------------------------------------------------------------
'A row opens a section of the report before its build and closes it
'after, so the parts of that build read as subsections under the row's
'own heading. The section is closed on every path out of the row, the
'failed one included.
'
'ONE HOST FOR THE RUN
'-------------------------------------------------------------------------------
'With a host on the instance path every row builds in the same hidden
'Excel over the same copy of the designer: the row's entries are written
'on Main here and cross with the row's first step, the file names of the
'row are checked before its build, and a row that fails has its build
'stopped through the host. With no host, or a host in place, the rows
'build in this Excel.
'@param lo ListObject. The T_Multi ListObject.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param bar ProgressBar. The bar over the rows. Nothing means no bar.
'@param builtCount Long. Answers the number of rows built and saved.
'@param failedCount Long. Answers the number of rows refused or failed.
'@param templatePath Optional String. The ribbon file every row is built with. Empty builds the buttons.
'@param host Optional GenerationHost. The host the rows build through, acquired with its copy open. Nothing builds here.
Public Sub GenerateMultipleRows(ByVal lo As ListObject, _
                                ByVal entry As DesignerEntry, _
                                ByVal bar As ProgressBar, _
                                ByRef builtCount As Long, _
                                ByRef failedCount As Long, _
                                Optional ByVal templatePath As String = vbNullString, _
                                Optional ByVal host As GenerationHost = Nothing)
    Dim rowIdx As Long
    Dim processed As Long
    Dim setupPath As String
    Dim outcomeText As String
    Dim rowBuilt As Boolean

    builtCount = 0
    failedCount = 0

    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIdx = 1 To lo.ListRows.Count
        setupPath = CellText(lo, rowIdx, COL_SETUPS)

        If LenB(setupPath) > 0 Then
            processed = processed + 1
            If Not bar Is Nothing Then
                bar.Update processed - 1, RowStatus(bar, processed, setupPath)
            End If

            EventsDesignerAdvanced.OpenLogSection _
                RowSectionTitle(lo, rowIdx, setupPath)

            FlushRowHeader lo, rowIdx, setupPath
            WriteRowEntries lo, rowIdx, entry, templatePath

            'The result cell shows the row's milestones while the row
            'builds and ends as the row's outcome below.
            rowBuilt = BuildRow(entry, outcomeText, RowCellRange(lo, rowIdx, COL_RESULT), host)

            If rowBuilt Then
                builtCount = builtCount + 1
                WriteRowCell lo, rowIdx, COL_OUTPUT_FILES, outcomeText
                WriteRowCell lo, rowIdx, COL_RESULT, RESULT_BUILT
            Else
                failedCount = failedCount + 1
                WriteRowCell lo, rowIdx, COL_RESULT, outcomeText
            End If

            FlushRowOutcome lo, rowIdx, outcomeText, rowBuilt
            EventsDesignerAdvanced.CloseLogSection

            If Not bar Is Nothing Then
                bar.Update processed, RowStatus(bar, processed, setupPath)
            End If
        ElseIf RowHasContent(lo, rowIdx) Then
            WriteRowCell lo, rowIdx, COL_RESULT, _
                         "Skipped: the " & COL_SETUPS & " cell is empty."
        End If
    Next rowIdx
End Sub

'@Description("The report section title of one row.")
'@details
'The row ID and the name of its setup file. The ID alone reads as
'"Operation- 3" in a report of ten sections and says nothing about which
'linelist it built.
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param setupPath String. The row's setup file path.
'@return String. The section title.
Public Function RowSectionTitle(ByVal lo As ListObject, _
                                ByVal rowIdx As Long, _
                                ByVal setupPath As String) As String
    Dim rowId As String

    rowId = CellText(lo, rowIdx, COL_ID)
    If LenB(rowId) = 0 Then rowId = "row " & CStr(rowIdx)

    RowSectionTitle = rowId & " - " & BaseName(setupPath)
End Function

'@Description("The global bar's status text for one row.")
'@param bar ProgressBar. The bar over the rows; its maximum is the row count.
'@param rowNumber Long. The position of the row among the rows the run builds.
'@param setupPath String. The row's setup file path.
'@return String. The status line, "linelist i of n - <setup file name>".
Private Function RowStatus(ByVal bar As ProgressBar, _
                           ByVal rowNumber As Long, _
                           ByVal setupPath As String) As String
    RowStatus = "linelist " & CStr(rowNumber) & " of " & CStr(bar.Maximum) & _
                " - " & BaseName(setupPath)
End Function

'@Description("Write one row's values into the Main entries through the entry manager.")
'@details
'Every mapped column is written, blanks included, so a row starts from
'its own values alone and behaves like a Main the user typed: an empty
'geobase is the optional entry, an empty epiweek reads as week 1, an
'empty required value is refused by the entry checks with the row's own
'message. The output files cell may carry the full path the last run
'wrote; OutputNameFromCell reduces it to the file name, which keeps a
're-run of the table stable.
'
'THE RIBBON FILE IS THE RUN'S, THE COLUMNS ARE THE ROW'S
'-------------------------------------------------------------------------------
'The template path arrives from the caller and is written like the mapped
'columns are, blanks included. The table carries no template column: the
'file is one for the whole run and the designer keeps it in its hidden
'names. Writing it on every row is what keeps the Main worksheet's own
'template cell out of the multi generation.
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param entry DesignerEntry. The entry manager over the Main worksheet.
'@param templatePath Optional String. The ribbon file of the run. Empty builds the buttons.
Public Sub WriteRowEntries(ByVal lo As ListObject, _
                           ByVal rowIdx As Long, _
                           ByVal entry As DesignerEntry, _
                           Optional ByVal templatePath As String = vbNullString)
    entry.AddInfo templatePath, "temppath"
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
'refused row answers False with the names of the entries that failed, or
'with the file-name clash the pre-flight found on the instance path; a
'failed row answers False with the error text, which names the kept
'__temp.xlsb when the build got far enough to keep one. The build has
'already closed that workbook, so the loop moves to the next row with
'nothing left open behind it. A failed row still asks for the abort
'through the route its steps took: a step that failed has already kept
'the file and the answer is a bare OK, and a build the instance stopped
'in has nothing to keep. The status target is the row's result cell: the
'build writes its milestones into it, and the caller overwrites it with
'the row's outcome.
'@param entry DesignerEntry. The entry manager, already loaded with the row.
'@param outcomeText String. Answers the written path on success and the fault text otherwise.
'@param statusTarget Range. One cell taking the build's milestone texts. Nothing means no writes.
'@param host GenerationHost. The host the build runs through, or Nothing.
'@return Boolean. True when the row built and saved.
Private Function BuildRow(ByVal entry As DesignerEntry, _
                          ByRef outcomeText As String, _
                          Optional ByVal statusTarget As Range = Nothing, _
                          Optional ByVal host As GenerationHost = Nothing) As Boolean
    Dim faults As Checking
    Dim clashText As String
    Dim keptPath As String

    On Error GoTo Fail

    If Not EventsDesignerAdvanced.ValidateEntries(entry, faults) Then
        'The names of the entries that failed, in the cell the user reads.
        'The line used to say only that the row was refused, which sent
        'somebody to the report to learn that a column of the row in front
        'of them was empty.
        outcomeText = "Refused: " & FaultList(faults)
        Exit Function
    End If

    If BuildsInInstance(host) Then
        clashText = EventsDesignerAdvanced.CheckBuildFileNames(host, entry, RowKeptPath(entry))
        If LenB(clashText) > 0 Then
            outcomeText = "Refused: " & Replace(clashText, vbLf, " ")
            Exit Function
        End If
    End If

    outcomeText = EventsDesignerAdvanced.GenerateOne(entry, statusTarget, host)
    BuildRow = True
    Exit Function

Fail:
    outcomeText = "Failed: " & Err.Description

    On Error Resume Next
    keptPath = EventsDesignerAdvanced.AbortBuild(host)
    If LenB(keptPath) > 0 Then
        If InStr(1, outcomeText, keptPath, vbTextCompare) = 0 Then
            outcomeText = outcomeText & " | kept " & keptPath
        End If
    End If
    On Error GoTo 0
End Function

'@Description("Whether the rows build in another instance through the host.")
'@param host GenerationHost. The host, or Nothing.
'@return Boolean. True when the host is on the instance path.
Private Function BuildsInInstance(ByVal host As GenerationHost) As Boolean
    If host Is Nothing Then Exit Function
    BuildsInInstance = Not host.InPlace
End Function

'@Description("The path a failed build of the row would keep its file at.")
'@param entry DesignerEntry. The entry manager, loaded with the row.
'@return String. <output folder>\OBTApp_\__temp.xlsb.
Private Function RowKeptPath(ByVal entry As DesignerEntry) As String
    Dim outputFolder As String

    outputFolder = entry.ValueOf("lldir")
    If Right$(outputFolder, 1) <> Application.PathSeparator Then
        outputFolder = outputFolder & Application.PathSeparator
    End If

    RowKeptPath = outputFolder & SCRATCH_FOLDER & Application.PathSeparator & KEPT_FILE_NAME
End Function

'@Description("Flush one log bundle naming the row before its build starts.")
'@details
'The first subsection of the row's section. The row's section is already
'open when this runs, so the bundle title reads as the subsection
'heading and the setup path lands in the entry under it.
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param setupPath String. The row's setup file path.
Private Sub FlushRowHeader(ByVal lo As ListObject, _
                           ByVal rowIdx As Long, _
                           ByVal setupPath As String)
    Dim rowChecks As Checking
    Dim rowId As String

    rowId = CellText(lo, rowIdx, COL_ID)
    If LenB(rowId) = 0 Then rowId = "row " & CStr(rowIdx)

    Set rowChecks = Checking.Create(TITLE_ROW_START)
    rowChecks.Add rowId, "Build started for: " & setupPath, checkingInfo

    EventsDesignerAdvanced.CollectIntoLog rowChecks
End Sub

'@Description("Flush one log bundle carrying the outcome of the row.")
'@details
'The last subsection of the row's section, so a reader who opens the
'report at a row learns how that row ended without going to the table.
'The written path is what a row that built answers; a row that failed
'answers the fault, and the entry carries the error scope so the
'severity filter of the report sheet finds it.
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param outcomeText String. The written path, or the fault text.
'@param rowBuilt Boolean. True when the row built and saved.
Private Sub FlushRowOutcome(ByVal lo As ListObject, _
                            ByVal rowIdx As Long, _
                            ByVal outcomeText As String, _
                            ByVal rowBuilt As Boolean)
    Dim rowChecks As Checking
    Dim rowId As String
    Dim scope As Byte
    Dim outcomeLabel As String

    rowId = CellText(lo, rowIdx, COL_ID)
    If LenB(rowId) = 0 Then rowId = "row " & CStr(rowIdx)

    If rowBuilt Then
        scope = checkingSuccess
        outcomeLabel = "Built: " & outcomeText
    Else
        scope = checkingError
        outcomeLabel = outcomeText
    End If

    Set rowChecks = Checking.Create(TITLE_ROW_OUTCOME)
    rowChecks.Add rowId, outcomeLabel, scope

    EventsDesignerAdvanced.CollectIntoLog rowChecks
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
    Dim cell As Range

    Set cell = RowCellRange(lo, rowIdx, colName)
    If cell Is Nothing Then Exit Function
    CellText = Trim$(CStr(cell.Value))
End Function

'@Description("The cell of a row by column header. A missing column answers Nothing.")
'@param lo ListObject. The T_Multi ListObject.
'@param rowIdx Long. The ListRows position of the row.
'@param colName String. The column header.
'@return Range. The one cell, or Nothing.
Private Function RowCellRange(ByVal lo As ListObject, _
                              ByVal rowIdx As Long, _
                              ByVal colName As String) As Range
    Dim col As ListColumn

    On Error Resume Next
    Set col = lo.ListColumns(colName)
    On Error GoTo 0

    If col Is Nothing Then Exit Function
    Set RowCellRange = lo.ListRows(rowIdx).Range.Cells(1, col.Index)
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
    Dim cell As Range

    Set cell = RowCellRange(lo, rowIdx, colName)
    If cell Is Nothing Then Exit Sub
    cell.Value = cellValue
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

'@Description("Put the designer's shared dropdowns on the columns that have no picker.")
'@details
'THE ENTRY CHECKS REFUSE A ROW MISSING ANY REQUIRED ENTRY, so a column
'with no way to fill it is a row that cannot build. Three of them were in
'that state: the interface language, the design and the epiweek start.
'The buttons of this group fill setups, geobases and output folders, and
'the dictionary language gets a list per row off its own setup file.
'
'The designer already registers all three lists for its own Main entries
'in DesignerPreparation, so the rows take the same ones. The validation
'goes on the whole column, which is why one call after a shape change
'covers every row the table now holds.
'
'The output file name stays typed by hand. It is the one required entry
'that is different per row and comes from nowhere but the user.
'@param lo ListObject. The T_Multi ListObject.
'@param drop DropdownLists. The dropdown manager of the host workbook.
Private Sub WireSharedColumns(ByVal lo As ListObject, ByVal drop As DropdownLists)
    Dim table As CustomTable

    If lo Is Nothing Then Exit Sub
    If drop Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set table = CustomTable.Create(lo)

    'A list the designer has never registered stops its own column alone,
    'and the other two still land.
    On Error Resume Next
    table.SetValidation COL_LANG_INTERFACE, drop, DROP_INTERFACE_LANGUAGES
    table.SetValidation COL_DESIGN, drop, DROP_DESIGN_VALUES
    table.SetValidation COL_EPIWEEK_START, drop, DROP_EPIWEEK_START
    Err.Clear
    On Error GoTo 0
End Sub

'@Description("True when the designer's GenerateMultiple worksheet is the one in front.")
'@details
'Every button of the Multi group works on the table of that one
'worksheet, and three of them write at the cursor row. A press made
'anywhere else has nothing to act on, so it does nothing at all: no
'dialog, no message, no write.
'
'A press from another workbook answers False too. A designer open beside
'a generated linelist shares one ribbon, and the group stays pressable
'over the other file.
'@return Boolean. True when the press may go on.
Private Function OnMultiSheet() As Boolean
    Dim current As Object

    If Application.ActiveWorkbook Is Nothing Then Exit Function
    If Not Application.ActiveWorkbook Is ThisWorkbook Then Exit Function

    Set current = Application.ActiveSheet
    If current Is Nothing Then Exit Function
    If Not TypeOf current Is Worksheet Then Exit Function

    OnMultiSheet = (StrComp(current.Name, SHEET_GENERATE_MULTIPLE, vbTextCompare) = 0)
End Function

'@Description("The entry names one row was refused on, in one line.")
'@details
'The keys of the entry checks ARE the entry names -- "setup path",
'"setup language", "design" -- so the row's result cell can name what to
'fill without anybody opening the report.
'@param faults Checking. What the entry checks filed.
'@return String. The names, comma separated.
Private Function FaultList(ByVal faults As Checking) As String
    Dim faultKeys As BetterArray
    Dim counter As Long

    If faults Is Nothing Then Exit Function

    Set faultKeys = faults.ListOfKeys
    If faultKeys Is Nothing Then Exit Function
    If faultKeys.Length = 0 Then Exit Function

    For counter = faultKeys.LowerBound To faultKeys.UpperBound
        If LenB(FaultList) > 0 Then FaultList = FaultList & ", "
        FaultList = FaultList & CStr(faultKeys.Item(counter))
    Next counter
End Function

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

    'For each setup file, extract languages and wire a per-row dropdown.
    'THE OPEN AND THE CLOSE ARE OWNED HERE, and the work between them is a
    'routine that raises nothing. A fault used to leave the setup workbook
    'sitting open on the screen: it reached the caller's handler, which has
    'no reference to close it with, and the message read "Unable to load
    'files" over a file the user could still see.
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
            WireRowLanguage lo, setupBook, filePath, currentRow, idCol, langCol, _
                            drop, skipped
            CloseQuietly setupBook
        End If

        currentRow = currentRow + 1
    Next idx
End Sub

'@Description("Build the language dropdown of one row from its own setup file.")
'@details
'Every fault is collected as a skip line and none is raised, because the
'caller holds the open setup workbook and is what closes it. A raise here
'would carry past the close.
'@param lo ListObject. The T_Multi ListObject.
'@param setupBook Workbook. The setup file, already open.
'@param filePath String. The path, for the skip lines.
'@param currentRow Long. Worksheet row number of the row being wired.
'@param idCol ListColumn. The ID column.
'@param langCol ListColumn. The dictionary language column.
'@param drop DropdownLists. The dropdown manager of the host workbook.
'@param skipped BetterArray. Collects one line per skipped step.
Private Sub WireRowLanguage(ByVal lo As ListObject, _
                            ByVal setupBook As Workbook, _
                            ByVal filePath As String, _
                            ByVal currentRow As Long, _
                            ByVal idCol As ListColumn, _
                            ByVal langCol As ListColumn, _
                            ByVal drop As DropdownLists, _
                            ByVal skipped As BetterArray)
    Dim tradSheet As Worksheet
    Dim langValues As BetterArray
    Dim rowId As String
    Dim dropName As String
    Dim langCell As Range

    On Error GoTo Failed

    On Error Resume Next
    Set tradSheet = setupBook.Worksheets(SHEET_TRANSLATIONS)
    On Error GoTo Failed

    If tradSheet Is Nothing Then
        skipped.Push "This setup file has no " & SHEET_TRANSLATIONS & _
                     " sheet: " & filePath
        Exit Sub
    End If

    Set langValues = EventsDesignerAdvanced.SetupLanguages(tradSheet)
    If langValues.Length = 0 Then
        skipped.Push "No language was found in: " & filePath
        Exit Sub
    End If

    'The dropdown is named after the row ID, in the shape Excel takes
    rowId = CStr(lo.Parent.Cells(currentRow, idCol.Range.Column).Value)
    dropName = SafeDropdownName(rowId)

    'Add or update the dropdown with extracted languages
    If drop.Exists(dropName) Then
        drop.Update langValues, dropName
    Else
        drop.Add langValues, dropName
    End If

    'Apply validation on the language cell using the dropdown
    Set langCell = lo.Parent.Cells(currentRow, langCol.Range.Column)
    drop.SetValidation langCell, dropName
    Exit Sub

Failed:
    skipped.Push "The language dropdown of row " & rowId & " was not built (" & _
                 filePath & "): " & Err.Description
End Sub

'@Description("The dropdown name of one row, in the shape Excel accepts.")
'@details
'The dropdown is named after the row ID, and an ID reads "Operation- 1".
'DropdownLists builds the workbook name dropdown_Operation-_1_lang out of
'that, and Excel refuses a name carrying a dash: "The syntax of this name
'isn't correct". The raise came back to the user as "Unable to load files"
'and it left the setup workbook open, so loading setups had never wired a
'single dropdown.
'
'Letters, digits and underscores are kept and every other character
'becomes an underscore, so one ID still answers one name and two rows
'still answer two.
'@param rowId String. The ID cell of the row.
'@return String. The dropdown name.
Private Function SafeDropdownName(ByVal rowId As String) As String
    Dim charIdx As Long
    Dim oneChar As String
    Dim cleanText As String

    For charIdx = 1 To Len(rowId)
        oneChar = Mid$(rowId, charIdx, 1)
        If oneChar Like "[A-Za-z0-9_]" Then
            cleanText = cleanText & oneChar
        Else
            cleanText = cleanText & "_"
        End If
    Next charIdx

    'A name starts with a letter or an underscore. An ID somebody typed as
    'a bare number would otherwise build one starting with a digit.
    If LenB(cleanText) = 0 Then cleanText = "row"
    If Left$(cleanText, 1) Like "[0-9]" Then cleanText = "_" & cleanText

    SafeDropdownName = cleanText & LANG_SUFFIX
End Function

'@Description("Close a workbook without saving and without raising.")
'@details
'Used on every path out of a step that opened a file. A close that raises
'would replace the fault the caller is about to report with its own.
'@param book Workbook. The workbook to close. Nothing is left alone.
Private Sub CloseQuietly(ByRef book As Workbook)
    If book Is Nothing Then Exit Sub

    On Error Resume Next
    book.Close saveChanges:=False
    Err.Clear
    On Error GoTo 0

    Set book = Nothing
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
