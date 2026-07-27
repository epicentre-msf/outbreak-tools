Attribute VB_Name = "OBTImport"
Attribute VB_Description = "Headless, AppleScript-callable wrappers around the Development manager: silent import and Codes-table rebuild"

Option Explicit

'@Folder("Rubberduck")
'@ModuleDescription("Headless, AppleScript-callable wrappers around the Development manager: silent import and Codes-table rebuild")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

' =============================================================================
' Headless import wrappers for the macOS AppleScript test loop.
'
' Thin layer AROUND the Development manager (Development.cls). OBTHeadless runs
' the suite; this module refreshes the workbook FROM src/ before the run, so the
' whole edit -> import -> run -> collect cycle happens with no manual upload.
'
'   OBTBuildCodeTables  -- rebuild the Codes-sheet import tables and the
'                          ModulesForTesting table from the registry
'                          intermediates in src/tests/.generated
'                          (code-tables.tsv + modules-for-testing.txt).
'   OBTSilentImport     -- Development.ImportAll with every prompt/alert off,
'                          so no dialog blocks the unattended run.
'
' Both are parameterless (run VB macro-callable): every path is derived in code
' from Root() = the opened workbook's OWN folder (ThisWorkbook.Path). run-tests.R
' assembles a self-contained run dir next to the workbook copy, so Excel reads
' the sources and manifest from the folder it opened the workbook from -- which
' the macOS sandbox auto-grants (reading/writing next to the open document needs
' no Full Disk Access; proven by the results CSV). This sidesteps the
' TCC prompt / -1712 hang a fixed staging path outside that folder triggers.
'
' Sub names carry no underscore on purpose: in a document/class module `Foo_Bar`
' parses as event `Bar` of object `Foo`, so the convention is kept everywhere.
'
' Run-dir layout the wrappers read (assembled by run-tests.R next to the copy):
'   Root()/classes/draft/*.cls        general classes the import pulls
'   Root()/tests/draft/*.bas          test modules + fixtures the import pulls
'   Root()/.generated/*               code-tables.tsv + modules-for-testing.txt
'
' Codes-sheet layout these wrappers own:
'   * folder path value cells (named ModulesCodes / ClassesImplementation /
'     TestsCodes) in column A, written by EnsureFolderRanges.
'   * ModulesForTesting rebuilt anchored at B2 (header) downwards.
'   * Development import tables rebuilt from E5 rightwards (its default).
' =============================================================================

Private Const CODE_SHEET As String = "Codes"
Private Const OUTPUT_SHEET As String = "testsOutputs"
Private Const MODULES_TABLE As String = "ModulesForTesting"
Private Const MODULES_ANCHOR As String = "B2"
Private Const GENERATED_DIR As String = ".generated"
Private Const CODE_TABLES_FILE As String = "code-tables.tsv"
Private Const MODULES_FILE As String = "modules-for-testing.txt"
Private Const LOG_FILE As String = "obt-import.log"

' Folder named ranges Development reads to resolve src/ paths.
Private Const MODULES_RANGE As String = "ModulesCodes"
Private Const CLASSES_RANGE As String = "ClassesImplementation"
Private Const TESTS_RANGE As String = "TestsCodes"

' Development table tags (mirror Development.cls; used to pick the create method).
Private Const TAG_CLASSES As String = "general classes"
Private Const TAG_MODULES As String = "general modules"
Private Const TAG_TESTS As String = "tests modules"
Private Const TAG_TEST_CLASSES As String = "tests classes"

' code-tables.tsv column order (produced by build-registry.R).
Private Const COL_FOLDER As Long = 0
Private Const COL_TAG As Long = 1
Private Const COL_COMPONENT As Long = 2
Private Const COL_INTERFACE As Long = 3


'@EntryPoint
'@sub-title Rebuild the Codes import tables and ModulesForTesting from the registry intermediates.
'@details
'Deletes every existing table on the Codes sheet, then rebuilds one Development
'import table per (folder, tag) group described in code-tables.tsv and the
'ModulesForTesting table from modules-for-testing.txt. Both files are produced
'by build-registry.R from src/tests/test-registry.yml. Runs dialog-free so it is
'safe under Apple Events automation.
Public Sub OBTBuildCodeTables()
    Dim prevAlerts As Boolean
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean

    prevAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo CleanExit

    LogReset
    LogStep "OBTBuildCodeTables: start (workbook path = " & ThisWorkbook.Path & ")"

    Dim codeSheet As Worksheet
    Dim outputSheet As Worksheet
    Set codeSheet = EnsureSheet(CODE_SHEET)
    LogStep "  Codes sheet resolved: " & codeSheet.Name
    Set outputSheet = EnsureSheet(OUTPUT_SHEET)
    LogStep "  testsOutputs sheet resolved: " & outputSheet.Name

    EnsureFolderRanges codeSheet
    LogStep "  folder ranges written: classes=" & StagePath("classes") & " tests=" & StagePath("tests")

    Dim manager As Development
    Set manager = Development.Create(codeSheet, codeSheet)
    manager.DisplayPrompts = False
    LogStep "  Development created"

    Dim generatedPath As String
    generatedPath = GeneratedFolder()
    LogStep "  generated dir: " & generatedPath

    ClearAllTables codeSheet
    LogStep "  tables cleared"

    BuildImportTables manager, JoinPath(generatedPath, CODE_TABLES_FILE)
    LogStep "  import tables built"

    BuildModulesForTesting codeSheet, JoinPath(generatedPath, MODULES_FILE)
    LogStep "  ModulesForTesting built"

CleanExit:
    If Err.Number <> 0 Then
        LogStep "  ERROR " & Err.Number & ": " & Err.Description
        Debug.Print "OBTBuildCodeTables error " & Err.Number & ": " & Err.Description
    Else
        LogStep "OBTBuildCodeTables: done"
    End If
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
End Sub

'@EntryPoint
'@sub-title Import every component listed in the Codes tables from src/, dialog-free.
'@details
'Drives Development.ImportAll with confirmation prompts and application alerts
'suppressed, so the automated import brings the registered classes and test
'modules into the workbook without any manual upload or dialog.
Public Sub OBTSilentImport()
    Dim prevAlerts As Boolean
    Dim prevEvents As Boolean
    Dim prevScreen As Boolean

    prevAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    prevScreen = Application.ScreenUpdating

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error GoTo CleanExit

    LogStep "OBTSilentImport: start"

    Dim codeSheet As Worksheet
    Set codeSheet = EnsureSheet(CODE_SHEET)
    LogStep "  Codes sheet resolved: " & codeSheet.Name

    EnsureFolderRanges codeSheet

    Dim manager As Development
    Set manager = Development.Create(codeSheet, codeSheet)
    manager.DisplayPrompts = False
    LogStep "  importing from " & StagePath("classes") & " and " & StagePath("tests")
    RemoveNonHeadlessModules
    manager.ImportAll
    LogStep "  ImportAll returned; components now = " & ComponentCount()

    LogCompileProbe

CleanExit:
    If Err.Number <> 0 Then
        LogStep "  ERROR " & Err.Number & ": " & Err.Description
        Debug.Print "OBTSilentImport error " & Err.Number & ": " & Err.Description
    Else
        LogStep "OBTSilentImport: done"
    End If
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = prevAlerts
End Sub


'@section Table Rebuild
'===============================================================================

'@sub-title Delete every ListObject on the sheet and clear the cells it occupied.
'@details
'Captures each table block (the tag and folder rows above the header down to the
'last data cell) before deleting, then clears those cells so a stale header can
'never offset the next horizontal table placement. Folder-range value cells in
'column A are never inside a table block, so they are preserved.
'@param sh Worksheet. The Codes worksheet to reset.
Private Sub ClearAllTables(ByVal sh As Worksheet)
    Dim addresses As Collection
    Dim lo As ListObject
    Dim item As Variant
    Dim topRow As Long
    Dim firstColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long

    Set addresses = New Collection

    For Each lo In sh.ListObjects
        firstColumn = lo.Range.Column
        lastColumn = lo.Range.Column + lo.Range.Columns.Count - 1
        lastRow = lo.Range.Row + lo.Range.Rows.Count - 1

        ' Extend two rows up for the tag/folder cells, clamped so a table anchored
        ' near the top (e.g. ModulesForTesting at row 2) never asks for row 0.
        topRow = lo.Range.Row - 2
        If topRow < 1 Then topRow = 1

        addresses.Add sh.Range(sh.Cells(topRow, firstColumn), _
                               sh.Cells(lastRow, lastColumn)).Address
    Next lo

    Do While sh.ListObjects.Count > 0
        sh.ListObjects(1).Delete
    Loop

    For Each item In addresses
        sh.Range(CStr(item)).ClearContents
    Next item
End Sub

'@sub-title Rebuild one Development import table per (folder, tag) group in code-tables.tsv.
'@param manager Development. The initialised manager whose table factories are used.
'@param tsvPath String. Absolute path of the generated code-tables.tsv file.
Private Sub BuildImportTables(ByVal manager As Development, ByVal tsvPath As String)
    Dim lines As Variant
    lines = SplitLines(ReadAllText(tsvPath))

    Dim orderedKeys As Collection
    Set orderedKeys = New Collection

    Dim rows As Collection
    Set rows = New Collection

    Dim idx As Long
    Dim fields As Variant
    Dim rawLine As String

    ' Parse every data line (skip the header row) into folder/tag/component/interface.
    For idx = LBound(lines) To UBound(lines)
        rawLine = CStr(lines(idx))
        If idx > LBound(lines) And LenB(Trim$(rawLine)) > 0 Then
            fields = Split(rawLine, vbTab)
            If UBound(fields) >= COL_COMPONENT Then
                rows.Add fields
                RememberKey orderedKeys, GroupKey(fields)
            End If
        End If
    Next idx

    Dim keyItem As Variant
    For Each keyItem In orderedKeys
        BuildGroupTable manager, rows, CStr(keyItem)
    Next keyItem
End Sub

'@sub-title Create and populate one import table for a single folder/tag group.
'@param manager Development. The initialised manager.
'@param rows Collection. Every parsed tsv row (arrays of fields).
'@param key String. The "folder|tag" group key to build.
Private Sub BuildGroupTable(ByVal manager As Development, ByVal rows As Collection, ByVal key As String)
    Dim parts As Variant
    parts = Split(key, "|")

    Dim folder As String
    Dim tag As String
    folder = parts(0)
    tag = parts(1)

    Dim lo As ListObject
    Dim isClass As Boolean
    Set lo = CreateTableForTag(manager, tag, isClass)
    If lo Is Nothing Then Exit Sub

    ' Folder cell sits two rows above the header (Development reads Cells(-1, 1)).
    lo.Range.Cells(-1, 1).Value = folder

    Dim names As Collection
    Dim flags As Collection
    Set names = New Collection
    Set flags = New Collection

    Dim item As Variant
    Dim fields As Variant
    For Each item In rows
        fields = item
        If StrComp(GroupKey(fields), key, vbTextCompare) = 0 Then
            names.Add CStr(fields(COL_COMPONENT))
            flags.Add InterfaceFlag(fields)
        End If
    Next item

    PopulateTable lo, names, flags, isClass
End Sub

'@fun-title Pick the Development factory for a tag and report whether it yields a class table.
'@param manager Development. The initialised manager.
'@param tag String. The scope tag from the registry.
'@param isClass Boolean. Set True when the created table carries the interface column.
'@return ListObject. The freshly created table, or Nothing for an unknown tag.
Private Function CreateTableForTag(ByVal manager As Development, ByVal tag As String, ByRef isClass As Boolean) As ListObject
    isClass = False

    Select Case LCase$(Trim$(tag))
        Case TAG_CLASSES
            isClass = True
            Set CreateTableForTag = manager.AddClassTable(False)
        Case TAG_TEST_CLASSES
            isClass = True
            Set CreateTableForTag = manager.AddClassTable(True)
        Case TAG_MODULES
            Set CreateTableForTag = manager.AddModuleTable(False)
        Case TAG_TESTS
            Set CreateTableForTag = manager.AddTestTable()
        Case Else
            ' Unknown tag: leave CreateTableForTag as Nothing so the caller skips it.
    End Select
End Function

'@sub-title Write component names (and interface flags for classes) into a table body.
'@param lo ListObject. The freshly created import table.
'@param names Collection. Component names in registry order.
'@param flags Collection. Yes/No interface flags aligned with names.
'@param isClass Boolean. True when the second column holds the interface flag.
Private Sub PopulateTable(ByVal lo As ListObject, ByVal names As Collection, ByVal flags As Collection, ByVal isClass As Boolean)
    Dim total As Long
    total = names.Count
    If total = 0 Then Exit Sub

    ' A newly created Development table has exactly one data row; grow/shrink to fit.
    Do While lo.ListRows.Count < total
        lo.ListRows.Add
    Loop
    Do While lo.ListRows.Count > total
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

    Dim rowIndex As Long
    For rowIndex = 1 To total
        lo.ListColumns(1).DataBodyRange.Cells(rowIndex, 1).Value = names.Item(rowIndex)
        If isClass And lo.ListColumns.Count >= 2 Then
            lo.ListColumns(2).DataBodyRange.Cells(rowIndex, 1).Value = flags.Item(rowIndex)
        End If
    Next rowIndex
End Sub

'@sub-title Rebuild the ModulesForTesting table the runner iterates, from modules-for-testing.txt.
'@param sh Worksheet. The Codes worksheet.
'@param txtPath String. Absolute path of the generated modules-for-testing.txt file.
Private Sub BuildModulesForTesting(ByVal sh As Worksheet, ByVal txtPath As String)
    Dim lines As Variant
    lines = SplitLines(ReadAllText(txtPath))

    Dim anchor As Range
    Set anchor = sh.Range(MODULES_ANCHOR)

    ' Clear a generous column block so an older, longer list never leaves orphans.
    sh.Range(anchor, anchor.Offset(1000, 0)).ClearContents

    anchor.Value = "modules"

    Dim written As Long
    Dim idx As Long
    Dim moduleName As String

    For idx = LBound(lines) To UBound(lines)
        moduleName = Trim$(CStr(lines(idx)))
        If LenB(moduleName) > 0 Then
            written = written + 1
            anchor.Offset(written, 0).Value = moduleName
        End If
    Next idx

    If written = 0 Then Exit Sub

    Dim tableRange As Range
    Set tableRange = sh.Range(anchor, anchor.Offset(written, 0))

    Dim lo As ListObject
    Set lo = sh.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    lo.Name = MODULES_TABLE
End Sub


'@section Grouping Helpers
'===============================================================================

'@fun-title Compose the "folder|tag" group key for a parsed tsv row.
Private Function GroupKey(ByVal fields As Variant) As String
    GroupKey = CStr(fields(COL_FOLDER)) & "|" & CStr(fields(COL_TAG))
End Function

'@fun-title Return the capitalised Yes/No interface flag for a class row.
Private Function InterfaceFlag(ByVal fields As Variant) As String
    Dim raw As String
    If UBound(fields) >= COL_INTERFACE Then raw = LCase$(Trim$(CStr(fields(COL_INTERFACE))))
    InterfaceFlag = IIf(raw = "yes", "Yes", "No")
End Function

'@sub-title Add a key to the ordered collection only the first time it is seen.
Private Sub RememberKey(ByVal keys As Collection, ByVal key As String)
    If Not HasKey(keys, key) Then keys.Add key, key
End Sub

'@fun-title Report whether a keyed collection already holds the given key.
Private Function HasKey(ByVal keys As Collection, ByVal key As String) As Boolean
    Dim probe As Variant
    On Error Resume Next
        probe = keys.Item(key)
        HasKey = (Err.Number = 0)
    On Error GoTo 0
End Function


'@section Path And File Helpers
'===============================================================================

'@sub-title Write the three src/ folder paths (in code) and point the named ranges at them.
'@details
'So the operator never wires the Dev sheet by hand: the ModulesCodes /
'ClassesImplementation / TestsCodes value cells are written in column A from
'the run root, and each workbook-scoped name is redefined to refer to its cell.
'Development.ImportAll then resolves the staging paths from these ranges.
'@param sh Worksheet. The Codes worksheet.
Private Sub EnsureFolderRanges(ByVal sh As Worksheet)
    SetFolderRange sh, sh.Range("A1"), MODULES_RANGE, StagePath("modules")
    SetFolderRange sh, sh.Range("A2"), CLASSES_RANGE, StagePath("classes")
    SetFolderRange sh, sh.Range("A3"), TESTS_RANGE, StagePath("tests")
End Sub

'@sub-title Write a folder path into a cell and (re)define a workbook name over it.
'@param sh Worksheet. The host worksheet.
'@param target Range. The single cell that holds the path value.
'@param rangeName String. The workbook-scoped name to redefine.
'@param pathValue String. The absolute folder path to store.
Private Sub SetFolderRange(ByVal sh As Worksheet, ByVal target As Range, ByVal rangeName As String, ByVal pathValue As String)
    target.Value = pathValue

    On Error Resume Next
        ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=rangeName, _
                           RefersTo:="='" & sh.Name & "'!" & target.Address
End Sub

'@fun-title The run root: the folder the workbook was opened from (sandbox auto-granted).
Private Function Root() As String
    Root = ThisWorkbook.Path
End Function

'@fun-title Build an absolute path to a subfolder of the run root.
Private Function StagePath(ByVal leaf As String) As String
    StagePath = JoinPath(Root(), leaf)
End Function

'@fun-title Absolute path of the run-root/.generated intermediates folder.
Private Function GeneratedFolder() As String
    GeneratedFolder = JoinPath(Root(), GENERATED_DIR)
End Function

'@fun-title Join a folder and a leaf with the platform path separator.
Private Function JoinPath(ByVal basePath As String, ByVal leaf As String) As String
    Dim sep As String
    sep = Application.PathSeparator

    If Right$(basePath, 1) = sep Then
        JoinPath = basePath & leaf
    Else
        JoinPath = basePath & sep & leaf
    End If
End Function

'@fun-title Read a whole text file as a single string (line-ending agnostic).
'@details
'Binary read so LF-only files written by R are captured intact; callers split on
'newlines themselves. Raises when the file is absent.
'@param path String. Absolute path of the file to read.
'@return String. The full file content.
Private Function ReadAllText(ByVal path As String) As String
    If Dir$(path) = vbNullString Then
        ThrowError "Generated file not found: " & path & " (run build-registry.R first)."
    End If

    Dim fileNum As Integer
    Dim content As String

    fileNum = FreeFile
    Open path For Binary Access Read As #fileNum
    content = Space$(LOF(fileNum))
    Get #fileNum, , content
    Close #fileNum

    ReadAllText = content
End Function

'@fun-title Split text into lines, normalising CRLF and CR to LF first.
Private Function SplitLines(ByVal text As String) As Variant
    Dim normalised As String
    normalised = Replace(text, vbCrLf, vbLf)
    normalised = Replace(normalised, vbCr, vbLf)
    SplitLines = Split(normalised, vbLf)
End Function

'@sub-title Raise an OBTImport-scoped error (surfaced by the caller's CleanExit log).
Private Sub ThrowError(ByVal message As String)
    Err.Raise vbObjectError + 2000, "OBTImport", message
End Sub


'@section Sheet And Diagnostics Helpers
'===============================================================================

'@fun-title Return a worksheet by name, creating it when the bootstrapped workbook lacks it.
'@details Self-heals the Codes/testsOutputs sheets so the loop never depends on
'the operator adding them by hand (parity with importing the probes in code).
'@param sheetName String. The worksheet to resolve or create.
'@return Worksheet. The existing or newly created worksheet.
Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim sh As Worksheet
    Dim candidate As Worksheet
    Dim total As Long

    On Error Resume Next
        total = ThisWorkbook.Worksheets.Count
    On Error GoTo 0
    LogStep "    EnsureSheet('" & sheetName & "'): Worksheets.Count = " & total

    ' Resolve by ITERATION, not Worksheets("name") string-indexing: under Apple
    ' Events automation on macOS that string lookup can fail to resolve an
    ' existing sheet (same fragility that bites worksheet-scoped Names), which
    ' would wrongly fall through to Add and raise error 91.
    For Each candidate In ThisWorkbook.Worksheets
        LogStep "      - found sheet: '" & candidate.Name & "'"
        If StrComp(candidate.Name, sheetName, vbTextCompare) = 0 Then
            Set sh = candidate
            Exit For
        End If
    Next candidate

    If sh Is Nothing Then
        LogStep "    EnsureSheet: creating missing sheet '" & sheetName & "'"
        Set sh = ThisWorkbook.Worksheets.Add
        sh.Name = sheetName
    End If

    Set EnsureSheet = sh
    LogStep "    EnsureSheet('" & sheetName & "') resolved OK"
End Function

'@fun-title Count VBA components currently in the project (post-import diagnostic).
Private Function ComponentCount() As String
    On Error Resume Next
        ComponentCount = CStr(ThisWorkbook.VBProject.VBComponents.Count)
        If Err.Number <> 0 Then ComponentCount = "VBE-ERR" & Err.Number
    On Error GoTo 0
End Function

'@sub-title Log an inventory of every component after import (safe post-import diagnostic).
'@details
'Runs IN-PROCESS at the tail of OBTSilentImport and records each component's
'name, type and line count to obt-import.log, which run-tests.R echoes on
'failure. This is what surfaces a stale or duplicate module -- e.g. an orphaned
'TestHelpers left in the saved workbook beside a freshly imported one -- whose
'colliding Public names would otherwise show up only as the opaque AppleScript
'-50 at OBTRunAllTests (a compile error names neither module nor symbol across
'the Apple Events boundary). Kept deliberately passive: it never invokes freshly
'imported code, because a compile error triggered mid-macro propagates as -50 and
'aborts the run rather than being trappable, so the inventory always completes.
'@vbext-ct component-type codes: 1 std module, 2 class, 3 MSForm, 100 document.
Private Sub LogCompileProbe()
    On Error Resume Next

    LogStep "=== COMPONENT INVENTORY (in-process, post-import) ==="

    Dim proj As Object
    Set proj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        LogStep "  VBProject inaccessible: Err " & Err.Number & " - " & Err.Description & _
                " (Trust access to the VBA project object model?)"
        Err.Clear
        LogStep "=== END INVENTORY ==="
        Exit Sub
    End If

    Dim comp As Object
    Dim lineCount As Long
    For Each comp In proj.VBComponents
        lineCount = -1
        lineCount = comp.CodeModule.CountOfLines
        LogStep "  comp: " & comp.Name & " | type=" & comp.Type & " | lines=" & lineCount
        Err.Clear
    Next comp

    LogStep "=== END INVENTORY ==="
    On Error GoTo 0
End Sub

'@sub-title Prune modules a headless Mac run must not carry, before import.
'@details
'One fault, swept in two places. It surfaces as the opaque AppleScript -50,
'because a compile failure names neither module nor symbol across Apple Events.
'  1. Orphan Test* standard modules absent from modules-for-testing.txt. Because
'     Development refreshes only REGISTERED names, a module left by a previous
'     registry lingers; once a newly registered module redefines the same Public
'     names, every unqualified call becomes "Ambiguous name detected" and the
'     whole project stops compiling. THIS is the real cause: a stale full
'     TestHelpers sat beside the registered TestHelpersLite, and both define
'     BusyApp and EnsureWorksheet.
'  2. CustomTestImplementation -- the ribbon/click runner. It calls BusyApp and
'     EnsureWorksheet unqualified, so it was a casualty of (1), not a cause. It
'     is dropped anyway: the headless loop runs through OBTHeadless and never
'     clicks a ribbon (see OBTHeadless header).
'NOT the cause: the clickRibbonTests(Control As IRibbonControl) signature. That
'was the first diagnosis and it is wrong -- IRibbonControl resolves fine on Mac.
'EventsLinelistRibbon uses it unguarded, and Linelist.cls copies that module into
'every generated linelist, which compiles on Mac. Do not re-add the claim.
'Registered test modules and all production/harness code are left untouched.
'Best-effort: guarded so a VBE-access failure can never break the import.
Private Sub RemoveNonHeadlessModules()
    On Error Resume Next

    Dim proj As Object
    Set proj = ThisWorkbook.VBProject
    If proj Is Nothing Then Exit Sub

    ' Registered-name set, framed with pipes for whole-name InStr matching.
    Dim registered As String
    registered = "|"

    Dim modulesFile As String
    modulesFile = JoinPath(GeneratedFolder(), MODULES_FILE)

    If Dir$(modulesFile) <> vbNullString Then
        Dim lines As Variant
        Dim idx As Long
        Dim moduleName As String
        lines = SplitLines(ReadAllText(modulesFile))
        For idx = LBound(lines) To UBound(lines)
            moduleName = Trim$(CStr(lines(idx)))
            If LenB(moduleName) > 0 Then registered = registered & moduleName & "|"
        Next idx
    End If

    ' Collect first, remove second: removing while iterating VBComponents is unsafe.
    Dim doomed As Collection
    Set doomed = New Collection

    Dim comp As Object
    For Each comp In proj.VBComponents
        If comp.Type = 1 Then                              ' vbext_ct_StdModule
            If StrComp(comp.Name, "CustomTestImplementation", vbTextCompare) = 0 Then
                doomed.Add comp.Name
            ElseIf StrComp(Left$(comp.Name, 4), "Test", vbTextCompare) = 0 Then
                If InStr(1, registered, "|" & comp.Name & "|", vbTextCompare) = 0 Then
                    doomed.Add comp.Name
                End If
            End If
        End If
    Next comp

    Dim item As Variant
    For Each item In doomed
        LogStep "  pruning module not used by the headless loop: " & CStr(item)
        proj.VBComponents.Remove proj.VBComponents(CStr(item))
    Next item

    On Error GoTo 0
End Sub

'@fun-title Absolute path of the diagnostics log, next to the running workbook.
Private Function LogPath() As String
    LogPath = ThisWorkbook.Path & Application.PathSeparator & LOG_FILE
End Function

'@sub-title Start a fresh diagnostics log for this run.
Private Sub LogReset()
    On Error Resume Next
        If Dir$(LogPath()) <> vbNullString Then Kill LogPath()
    On Error GoTo 0
End Sub

'@sub-title Append one diagnostics line to the log (best-effort; never raises).
'@param message String. The line to record.
Private Sub LogStep(ByVal message As String)
    Dim fileNum As Integer
    On Error Resume Next
        fileNum = FreeFile
        Open LogPath() For Append As #fileNum
        Print #fileNum, message
        Close #fileNum
    On Error GoTo 0
End Sub
