VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist data export operations"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("Linelist data export operations")

Option Explicit

'@class LLExporter
'@description
'Exports the data of a linelist into standalone workbooks. Four modes are
'offered: ExportCustom writes one of the exports defined on the export sheet,
'ExportMigration writes the whole linelist for a move to another workbook,
'ExportAnalysis writes the analysis sheets, and ExportGeo writes the geobase.
'Each mode builds a temporary output workbook, fills it with the data, the
'metadata and the formatting the mode calls for, then saves it and closes it.
'CloseAll closes whatever is still open. Create binds the exporter to a
'workbook the caller holds; CreateFromFile opens a linelist file itself, and
'the workbook it opened is closed by CloseAll.
'
'WHAT AN EXPORT CARRIES
'-------------------------------------------------------------------------------
'ExportOptionsFor and MigrationOptions build the options line the metadata
'sheet records, ExcludedVariablesFor answers the variables a mode leaves out,
'EpiWeekStart answers the first day of the epidemiological week, and
'LastExportPassword answers the password the last export was saved under.
'LastFailure answers what the last export failed on, for a caller writing the
'failure into the log.
'@depends LLdictionary, LLVariables, LLExport, Passwords, CustomTable, LLLog
'@depends DataSheet, AnaTabIds, LLGeo, LLChoices, HiddenNames, LLTranslation
'@depends ShowHide, ShowHideLayout, ShowHideStore, BetterArray, Checking
'@depends DropdownLists, ChoiceKeys


'@section Enumerations
'===============================================================================
'@description Scopes for export operations, used to tag metadata in output workbooks.

Public Enum ExportScope
    ExportScopeCustom = 1
    ExportScopeMigration = 2
    ExportScopeAnalysis = 3
End Enum


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "LLExporter"

' Source workbook sheet names
Private Const DICTIONARYSHEET As String = "Dictionary"
Private Const CHOICESSHEET As String = "Choices"
Private Const DICTTRANSLATIONSHEET As String = "Translations"
Private Const LINELISTTRANSLATIONSHEET As String = "LinelistTranslation"
Private Const METADATASHEET As String = "Metadata"
Private Const PASSWORDSHEET As String = "__pass"
Private Const TEMPSHEET As String = "__temp"
Private Const EXPORTSHEET As String = "Exports"
Private Const ANANAMESSHEET As String = "__ana_tabnames"
Private Const SHOWHIDESHEET As String = "__show_hide"
Private Const GEOSHEET As String = "Geo"

' The names a linelist generated before the internal-sheet rename gave the
' temp and show/hide sheets. CreateFromFile opens such files, so the reads
' on the source workbook look for the new name first and fall back to these.
Private Const TEMPSHEET_OLD As String = "temp__"
Private Const SHOWHIDESHEET_OLD As String = "show_hide__"

' Prefixes of the two companion sheets a data entry sheet can have
Private Const PRINTPREFIX As String = "print_"
Private Const CRFPREFIX As String = "crf_"

' The three values the "header format" column of the Exports sheet can hold.
' An empty cell reads as HEADER_NAMES.
Private Const HEADER_NAMES As String = "variable names"
Private Const HEADER_LABELS As String = "variables labels"
Private Const HEADER_NAMES_LABELS As String = "variable names + variable labels"

' The serialised export options row, and the separators it is read back with
Private Const OPTION_SEPARATOR As String = " | "
Private Const OPTION_PAIR_SEPARATOR As String = ": "

' P-codes tag list for removal from custom exports
Private Const PCODES_GEO_PCODE_ADM1 As String = "geo_pcode_adm1"
Private Const PCODES_GEO_PCODE_ADM2 As String = "geo_pcode_adm2"
Private Const PCODES_GEO_PCODE_ADM3 As String = "geo_pcode_adm3"
Private Const PCODES_GEO_PCODE_ADM4 As String = "geo_pcode_adm4"
Private Const PCODES_GEO_CONCAT_ADM1 As String = "geo_concat_adm1"
Private Const PCODES_GEO_CONCAT_ADM2 As String = "geo_concat_adm2"
Private Const PCODES_GEO_CONCAT_ADM3 As String = "geo_concat_adm3"
Private Const PCODES_GEO_CONCAT_ADM4 As String = "geo_concat_adm4"

' How the "admin levels" cell of an Exports row is written, and what an empty
' one is called on the options line a file carries about itself.
Private Const ADMIN_LEVEL_SEPARATOR As String = "+"
Private Const ADMIN_LEVEL_PREFIX As String = "admin"
Private Const ADMIN_LEVEL_COUNT As Long = 4
Private Const ADMIN_LEVELS_ALL As String = "all"


'@section State
'===============================================================================

Private Type TLLExporter
    sourceWkb As Workbook
    lastPassword As String
    lastFailure As String
    activeWkb As Workbook
    sealed As Boolean
    openedSource As Boolean
    passObj As Passwords
    triedPass As Boolean
    userLog As LLLog
    triedLog As Boolean
End Type

Private this As TLLExporter


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create an exporter from the source linelist workbook
'@details
'Factory method that validates the workbook is not Nothing, then constructs
'and returns a new LLExporter instance. Called on the predeclared instance.
'@param sourceWkb Workbook. The linelist workbook containing the data to export.
'@return LLExporter. A fully initialised instance ready for export operations.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook) As LLExporter

    If sourceWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "LLExporter requires a source workbook"
    End If

    Dim newExporter As LLExporter
    Set newExporter = New LLExporter
    Set newExporter.SourceWorkbook = sourceWkb
    newExporter.Seal
    Set Create = newExporter
End Function

'@label:create-from-file
'@sub-title Create an exporter on a linelist file, opened here
'@details
'Opens the linelist file at the given path, read-only and with the given
'password, and builds an exporter on it. The workbook opened here belongs to
'this instance, and CloseAll is what closes it: the export walks write into the
'source's temp sheet, so the copy in memory changes, and a read-only open with
'no save is how none of that reaches the file. When the file is already open in
'this Excel session, the open workbook is used as it stands and stays out of
'CloseAll's reach.
'
'The caller decides the application state. A linelist carries open events, so
'an open with events running starts them; the form logic applies its busy
'state, events suppressed, before calling here.
'@param filePath String. Full path of the linelist file.
'@param password Optional String. The password the file opens with. Defaults to none.
'@return LLExporter. An instance bound to the opened workbook.
'@throws ProjectError.ElementNotFound When no file sits at the given path.
'@throws ProjectError.ErrorUnexpectedState When the file does not open with the
'  given password, or when the opened workbook misses a linelist sheet.
'@export
Public Function CreateFromFile(ByVal filePath As String, _
                               Optional ByVal password As String = vbNullString) As LLExporter

    Dim wkb As Workbook
    Dim missingSheet As String
    Dim openedHere As Boolean

    If LenB(Dir(filePath)) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "There is no file at " & filePath
    End If

    Set wkb = FindOpenWorkbook(filePath)

    If wkb Is Nothing Then
        'A wrong password or an unreadable file raises inside Workbooks.Open,
        'and the raise below is what carries the failure to the caller.
        On Error Resume Next
        Set wkb = Workbooks.Open(fileName:=filePath, ReadOnly:=True, password:=password)
        On Error GoTo 0
        openedHere = True
    End If

    If wkb Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "The file did not open with the given password: " & filePath
    End If

    missingSheet = MissingLinelistSheet(wkb)
    If LenB(missingSheet) > 0 Then
        If openedHere Then wkb.Close savechanges:=False
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "The workbook is not a linelist, it has no " & missingSheet & " sheet"
    End If

    Dim newExporter As LLExporter
    Set newExporter = New LLExporter
    Set newExporter.SourceWorkbook = wkb
    If openedHere Then newExporter.MarkSourceOpened
    newExporter.Seal
    Set CreateFromFile = newExporter
End Function

'@label:find-open-workbook
'@sub-title The workbook of this session already open on a path, or Nothing
'@param filePath String. Full path to look for.
'@return Workbook. The open workbook, or Nothing when the path is open nowhere.
Private Function FindOpenWorkbook(ByVal filePath As String) As Workbook
    Dim wkb As Workbook

    For Each wkb In Application.Workbooks
        If LCase$(wkb.FullName) = LCase$(filePath) Then
            Set FindOpenWorkbook = wkb
            Exit Function
        End If
    Next
End Function

'@label:missing-linelist-sheet
'@sub-title The first sheet a linelist carries and a workbook misses
'@details
'The export walks read the dictionary, the export definitions, the passwords,
'the geobase and the temp sheet, so those five names are what marks a workbook
'as a linelist here.
'@param wkb Workbook. The workbook to look through.
'@return String. The first missing sheet name, or an empty string when all five are there.
Private Function MissingLinelistSheet(ByVal wkb As Workbook) As String
    Dim requiredSheets As BetterArray
    Dim counter As Long
    Dim sheetName As String
    Dim foundSheet As Worksheet

    Set requiredSheets = New BetterArray
    requiredSheets.Push DICTIONARYSHEET, EXPORTSHEET, PASSWORDSHEET, GEOSHEET

    For counter = requiredSheets.LowerBound To requiredSheets.UpperBound
        sheetName = requiredSheets.Item(counter)
        Set foundSheet = Nothing
        'A missing sheet raises on the read, and Nothing is the answer read back.
        On Error Resume Next
        Set foundSheet = wkb.Worksheets(sheetName)
        On Error GoTo 0
        If foundSheet Is Nothing Then
            MissingLinelistSheet = sheetName
            Exit Function
        End If
    Next

    'The temp sheet answers to either name, since the file may be older than
    'the internal-sheet rename.
    If InternalSheetOf(wkb, TEMPSHEET, TEMPSHEET_OLD) Is Nothing Then
        MissingLinelistSheet = TEMPSHEET
    End If
End Function

'@label:internal-sheet-of
'@sub-title An internal worksheet of a workbook under either of its names
'@details
'A linelist generated before the internal-sheet rename carries the old
'trailing name, so the new name is looked for first and the old one answers
'when it is absent.
'@param wkb Workbook. The workbook to look through.
'@param newName String. The name the sheet carries today.
'@param oldName String. The name an older linelist gave it.
'@return Worksheet. The worksheet, or Nothing when the workbook has neither name.
Private Function InternalSheetOf(ByVal wkb As Workbook, _
                                 ByVal newName As String, _
                                 ByVal oldName As String) As Worksheet
    Dim sh As Worksheet

    'A missing sheet raises on the read, and Nothing is the answer read back.
    On Error Resume Next
    Set sh = wkb.Worksheets(newName)
    If sh Is Nothing Then Set sh = wkb.Worksheets(oldName)
    On Error GoTo 0

    Set InternalSheetOf = sh
End Function


'@section Public Accessors
'===============================================================================

'@label:source-workbook
'@prop-title Source linelist workbook
'@return Workbook. The linelist workbook bound to this exporter.
Public Property Get SourceWorkbook() As Workbook
    Set SourceWorkbook = this.sourceWkb
End Property

'@label:source-workbook-set
'@prop-title Assign the source linelist workbook
'@param wb Workbook. The workbook to store.
Public Property Set SourceWorkbook(ByVal wb As Workbook)
    GuardNotSealed "SourceWorkbook"
    Set this.sourceWkb = wb
End Property

'@label:seal
'@prop-title Prevent further changes to setup-only properties
Public Sub Seal()
    this.sealed = True
End Sub

'@label:opened-from-file
'@prop-title Whether this instance opened its source workbook itself
'@details
'True on an instance built by CreateFromFile that opened the file, and that is
'the instance whose CloseAll closes the source. An instance built on a workbook
'the caller already held answers False and leaves that workbook open.
'@return Boolean. True when CloseAll closes the source workbook.
'@export
Public Property Get OpenedFromFile() As Boolean
    OpenedFromFile = this.openedSource
End Property

'@label:mark-source-opened
'@sub-title Record that this instance opened its source workbook
'@details
'Set at creation only, by CreateFromFile, before the seal.
Public Sub MarkSourceOpened()
    GuardNotSealed "MarkSourceOpened"
    this.openedSource = True
End Sub

'@label:guard-not-sealed
'@sub-title Raise when a setup-only property is written after sealing
'@param propName String. The name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:last-export-password
'@prop-title Password used in the last export operation
'@details
'Returns the password applied to the most recent custom export, or an
'empty string when no password was used. Modules display this value to
'the user after a successful export so they can share it securely.
'@return String. The password, or vbNullString.
'@export
Public Property Get LastExportPassword() As String
    LastExportPassword = this.lastPassword
End Property

'@label:last-failure
'@prop-title What the last export failed on, in the words of the error itself
'@details
'An export that fails raises, and the caller reads a description that names
'the method and nothing else, because the error crosses out of this class.
'The number, the source and the text are kept here before the raise, so the
'caller can log what actually refused.
'
'Every export empties this on entry, so a value here belongs to the export
'that just ran. It is empty after an export that worked.
'@return String. The error number, source and text, or vbNullString.
'@export
Public Property Get LastFailure() As String
    LastFailure = this.lastFailure
End Property

'@label:epi-week-start
'@prop-title First day of the epidemiological week, as the linelist holds it
'@details
'RNG_EpiWeekStart is a HiddenNames entry and its definition is a quoted string,
'so the name has no range behind it. Every reader that went through
'Names("RNG_EpiWeekStart").RefersToRange raised 1004 into a swallow and answered
'an empty string, which is why the epiweek tag of every export file was blank.
'The store is what answers it, and this is what AddMetadataTags writes.
'@return String. The stored value, empty when the workbook carries no such name.
'@export
Public Function EpiWeekStart() As String
    EpiWeekStart = HiddenNames.Create(this.sourceWkb).ValueAsString("RNG_EpiWeekStart")
End Function

'@label:export-options-for
'@sub-title The options line a numbered export would write into its file
'@details
'The third read seam of this class, beside EpiWeekStart and
'ExcludedVariablesFor. It answers exactly what AddMetadataTags writes into the
'`export_options` row, so what a file will say about itself can be read without
'writing a file.
'
'The line is what lets an import know whether a file can be read back at all.
'`variable names: no` marks a file whose header row carries the labels a user
'reads and no variable name anywhere, and `header row: 2` marks one carrying
'both, names on the second row.
'@param exportNumber Long. Which export definition to read (1-based).
'@param useFilters Optional Boolean. Whether the data would come from the
'  filtered companions. Defaults to False.
'@return String. The options of that export on one line.
'@export
Public Function ExportOptionsFor(ByVal exportNumber As Long, _
                                 Optional ByVal useFilters As Boolean = False) As String
    ExportOptionsFor = ExportOptionsTag(ExportScopeCustom, exportNumber, useFilters)
End Function

'@label:migration-options
'@sub-title The options line an export for migration writes into its file
'@return String. The options of a migration export on one line.
'@export
Public Function MigrationOptions() As String
    MigrationOptions = ExportOptionsTag(ExportScopeMigration, 0, False)
End Function

'@label:excluded-variables-for
'@sub-title The variables a numbered export leaves out
'@details
'The selection an export makes before it writes anything: the geo variables of
'every admin level the "admin levels" cell leaves out, the geo p-code and geo
'concat variables when the export row turns p-codes off, and the personal
'identifiers when it turns those off. Reading it costs four column reads at most
'and writes nothing, so a caller can ask what a file will be missing before
'asking for the file.
'@param exportNumber Long. Which export definition to read (1-based).
'@return BetterArray. The variable names the export drops.
'@export
Public Function ExcludedVariablesFor(ByVal exportNumber As Long) As BetterArray
    Dim expObj As LLExport
    Dim nbExp As Long

    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()
    Set ExcludedVariablesFor = ExcludedVariables(SourceDictionary(nbExp), _
                                                 expObj, exportNumber)
End Function


'@section Export Methods
'===============================================================================

'@label:export-custom
'@sub-title Export a numbered custom export definition
'@details
'Creates an output workbook containing filtered data, optional metadata
'sheets, and optional analysis sheets according to the export definition
'identified by exportNumber on the Exports sheet. Removes p-codes and
'personal identifiers if configured, prepares a temp dictionary, sorts it,
'then populates the output workbook with data sheets and formatting.
'Saves the workbook with an optional password and returns the file path.
'@param exportNumber Long. Which export definition to use (1-based).
'@param folderPath String. Target folder for the saved file.
'@param useFilters Optional Boolean. When True, exports filtered HList data. Defaults to False.
'@return String. Saved file path, or vbNullString on failure.
'@export
Public Function ExportCustom(ByVal exportNumber As Long, _
                             ByVal folderPath As String, _
                             Optional ByVal useFilters As Boolean = False) As String

    Dim outwb As Workbook
    Dim expObj As LLExport
    Dim dict As LLdictionary
    Dim pass As Passwords
    Dim expDict As LLdictionary
    Dim sheetNames As BetterArray
    Dim excluded As BetterArray
    Dim expName As String
    Dim withAnalysis As Boolean
    Dim withMetadata As Boolean
    Dim useLabels As Boolean
    Dim startLine As Long
    Dim headerFormat As String
    Dim fileName As String
    Dim filePath As String
    Dim fileExt As String
    Dim fileFormat As Long
    Dim addPassword As Boolean
    Dim counter As Long
    Dim prevSheetName As String
    Dim firstSheetName As String
    Dim nbExp As Long
    Dim atStep As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    ExportCustom = vbNullString
    this.lastPassword = vbNullString
    this.lastFailure = vbNullString

    On Error GoTo RecordFailure

    atStep = "reading the Exports sheet"
    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()

    atStep = "reading the dictionary"
    Set dict = SourceDictionary(nbExp)

    atStep = "reading the passwords"
    Set pass = PasswordsObject()
    expName = "export " & CStr(exportNumber)

    ' The p-codes and the personal identifiers this export leaves out. The
    ' source dictionary is read and never written.
    Set excluded = ExcludedVariables(dict, expObj, exportNumber)

    ' Filter dictionary for this export and write to temp sheet
    PrepareTempDictionary dict, expName, False, excluded
    Set expDict = ExportDictionaryFromTemp(nbExp)

    ' Sort export dictionary by table name and export column
    SortExportDictionary expDict, expName

    ' Create output workbook
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add metadata if configured. One flag decides the whole metadata set:
    ' "export metadata sheets = yes" ships the Metadata, Choices, Translations
    ' and Dictionary worksheets AND the show/hide sheet, so the tags the file
    ' carries describe what is really in it. Owner decision, 2026-08-01.
    ' AddShowHide used to run from ExportMigration alone while the tag was
    ' written True here, so every custom export promised a __show_hide sheet it
    ' had no copy of.
    withMetadata = (expObj.ColumnValue(exportNumber, "export metadata sheets") = "yes")
    If withMetadata Then
        AddMetadataSheets outwb, dict, nbExp
        AddCustomDropdowns outwb
        AddDictionary outwb, dict, nbExp, expName, excluded
        AddMetadataTags outwb, ExportScopeCustom, exportNumber, _
                        includeShowHide:=True, keepLabels:=True, _
                        useFilters:=useFilters
        AddShowHide outwb
        prevSheetName = METADATASHEET
    Else
        prevSheetName = firstSheetName
    End If

    ' Add data worksheets
    Set sheetNames = expDict.UniqueValues("sheet name")
    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        outwb.Worksheets.Add(After:=outwb.Worksheets(prevSheetName)).Name = _
            sheetNames.Item(counter)
        prevSheetName = sheetNames.Item(counter)
    Next counter

    ' Add analysis if configured
    withAnalysis = (expObj.ColumnValue(exportNumber, "export analyses sheets") = "yes")
    If withAnalysis Then AddAnalysisSheets outwb, prevSheetName

    ' Delete the placeholder first sheet. The handler is re-armed rather than
    ' cleared, so the rest of the walk still records what it refused on.
    On Error Resume Next
    outwb.Worksheets(firstSheetName).Delete
    On Error GoTo RecordFailure

    ' Export data to each sheet
    atStep = "writing the data"
    headerFormat = expObj.ColumnValue(exportNumber, "header format")
    useLabels = (headerFormat = HEADER_LABELS) Or _
                (headerFormat = HEADER_NAMES_LABELS)
    startLine = IIf(headerFormat = HEADER_NAMES_LABELS, 2, 1)

    AddDataSheets outwb, expDict, useLabels, startLine, useFilters

    ' Build file path and save
    fileName = expObj.ExportFileName(exportNumber, dict, pass)
    fileExt = expObj.ColumnValue(exportNumber, "file format")
    If LenB(fileExt) = 0 Then fileExt = "xlsx"
    fileFormat = IIf(fileExt = "xlsb", xlExcel12, xlOpenXMLWorkbook)
    filePath = BuildFilePath(folderPath, fileName, fileExt)

    'No passwords object means no private key, so a row asking for a password
    'gets a file without one. PasswordsObject has already logged why.
    addPassword = (expObj.ColumnValue(exportNumber, "password") = "yes") And _
                  Not (pass Is Nothing)
    If addPassword Then this.lastPassword = pass.Value("privatekey")

    atStep = "saving " & filePath
    SaveWorkbook outwb, filePath, fileFormat, this.lastPassword
    ExportCustom = filePath
    Exit Function

RecordFailure:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    RecordExportFailure "custom export " & CStr(exportNumber), atStep, _
                        errNumber, errSource, errMessage
    Err.Raise errNumber, errSource, errMessage
End Function

'@label:export-migration
'@sub-title Export all data for migration to another linelist
'@details
'Creates an output workbook with full unfiltered data, metadata sheets,
'dictionary, choices, translations, and metadata tags for re-import.
'Optionally includes show/hide state with column widths. No password is
'applied. Returns the saved file path.
'@param folderPath String. Target folder for the saved file.
'@param includeShowHide Optional Boolean. When True, includes show/hide state. Defaults to False.
'@param keepLabels Optional Boolean. When True, marks editable labels for update on import. Defaults to False.
'@return String. Saved file path, or vbNullString on failure.
'@export
Public Function ExportMigration(ByVal folderPath As String, _
                                Optional ByVal includeShowHide As Boolean = False, _
                                Optional ByVal keepLabels As Boolean = False) As String

    Dim outwb As Workbook
    Dim expObj As LLExport
    Dim dict As LLdictionary
    Dim pass As Passwords
    Dim expDict As LLdictionary
    Dim sheetNames As BetterArray
    Dim excluded As BetterArray
    Dim fileName As String
    Dim filePath As String
    Dim counter As Long
    Dim prevSheetName As String
    Dim firstSheetName As String
    Dim nbExp As Long
    Dim atStep As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    ExportMigration = vbNullString
    this.lastPassword = vbNullString
    this.lastFailure = vbNullString

    On Error GoTo RecordFailure

    atStep = "reading the Exports sheet"
    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()

    atStep = "reading the dictionary"
    Set dict = SourceDictionary(nbExp)

    atStep = "reading the passwords"
    Set pass = PasswordsObject()

    ' A migration carries the whole dictionary, so it excludes nothing.
    Set excluded = New BetterArray

    ' Write full dictionary to temp sheet (no filtering)
    atStep = "writing the dictionary to the temp sheet"
    PrepareTempDictionary dict, vbNullString, True, excluded

    atStep = "reading the dictionary back off the temp sheet"
    Set expDict = ExportDictionaryFromTemp(nbExp)

    ' Create output workbook
    atStep = "adding the output workbook"
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add metadata sheets
    atStep = "adding the metadata sheets"
    AddMetadataSheets outwb, dict, nbExp

    atStep = "adding the custom dropdowns"
    AddCustomDropdowns outwb

    atStep = "adding the dictionary"
    AddDictionary outwb, dict, nbExp, "__all__", excluded

    atStep = "adding the metadata tags"
    AddMetadataTags outwb, ExportScopeMigration, 0, includeShowHide, keepLabels, False

    ' Add show/hide with column widths if requested
    If includeShowHide Then
        atStep = "adding the show/hide choices"
        AddShowHide outwb
    End If

    prevSheetName = METADATASHEET

    ' Add data worksheets
    atStep = "adding one worksheet per data sheet"
    Set sheetNames = expDict.UniqueValues("sheet name")
    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        outwb.Worksheets.Add(After:=outwb.Worksheets(prevSheetName)).Name = _
            sheetNames.Item(counter)
        prevSheetName = sheetNames.Item(counter)
    Next counter

    ' Delete the placeholder first sheet. The handler is re-armed rather than
    ' cleared: On Error GoTo 0 here would leave the rest of the walk with no
    ' handler at all, and the account below would never be written.
    On Error Resume Next
    outwb.Worksheets(firstSheetName).Delete
    On Error GoTo RecordFailure

    ' Export all data (no filtering, no labels)
    atStep = "writing the data"
    AddDataSheets outwb, expDict, False, 1, False

    ' Save as xlsx (no password for migration)
    atStep = "building the file name"
    fileName = expObj.ExportFileName(0, dict, pass, exportAll:=True)
    filePath = BuildFilePath(folderPath, fileName, "xlsx")

    atStep = "saving " & filePath
    SaveWorkbook outwb, filePath, xlOpenXMLWorkbook, vbNullString

    ExportMigration = filePath
    Exit Function

RecordFailure:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    RecordExportFailure "migration export", atStep, errNumber, errSource, errMessage
    Err.Raise errNumber, errSource, errMessage
End Function

'@label:export-analysis
'@sub-title Export analysis worksheets only
'@details
'Creates an output workbook containing only the four analysis sheets
'(univariate, time series, spatial, spatio-temporal) with their values,
'formatting, and graphs. No data or metadata sheets are included. Saves
'as xlsx without password.
'@param folderPath String. Target folder for the saved file.
'@return String. Saved file path, or vbNullString on failure.
'@export
Public Function ExportAnalysis(ByVal folderPath As String) As String

    Dim outwb As Workbook
    Dim expObj As LLExport
    Dim dict As LLdictionary
    Dim pass As Passwords
    Dim fileName As String
    Dim filePath As String
    Dim firstSheetName As String
    Dim nbExp As Long
    Dim atStep As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    ExportAnalysis = vbNullString
    this.lastPassword = vbNullString
    this.lastFailure = vbNullString

    On Error GoTo RecordFailure

    atStep = "reading the Exports sheet"
    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()

    atStep = "reading the dictionary"
    Set dict = SourceDictionary(nbExp)

    atStep = "reading the passwords"
    Set pass = PasswordsObject()

    ' Create output workbook
    atStep = "adding the output workbook"
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add analysis sheets
    atStep = "adding the analysis sheets"
    AddAnalysisSheets outwb, firstSheetName

    ' Delete the placeholder first sheet. The handler is re-armed rather than
    ' cleared, so the rest of the walk still records what it refused on.
    On Error Resume Next
    outwb.Worksheets(firstSheetName).Delete
    On Error GoTo RecordFailure

    ' Save
    atStep = "building the file name"
    fileName = expObj.ExportFileName(0, dict, pass, exportAll:=True)
    filePath = BuildFilePath(folderPath, fileName & "_analysis", "xlsx")

    atStep = "saving " & filePath
    SaveWorkbook outwb, filePath, xlOpenXMLWorkbook, vbNullString

    ExportAnalysis = filePath
    Exit Function

RecordFailure:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    RecordExportFailure "analysis export", atStep, errNumber, errSource, errMessage
    Err.Raise errNumber, errSource, errMessage
End Function

'@label:export-geo
'@sub-title Export geobase data to a separate workbook
'@details
'Creates an output workbook and delegates to LLGeo.ExportToFile to write
'the geobase data. When onlyHistoric is True, exports only the historic
'subset. The file name includes a date stamp and optionally a "_historic"
'suffix.
'@param folderPath String. Target folder for the saved file.
'@param onlyHistoric Optional Boolean. When True, exports only the historic geobase. Defaults to False.
'@return String. Saved file path, or vbNullString on failure.
'@export
Public Function ExportGeo(ByVal folderPath As String, _
                           Optional ByVal onlyHistoric As Boolean = False) As String

    Dim outwb As Workbook
    Dim geo As LLGeo
    Dim geobaseName As String
    Dim filePath As String
    Dim atStep As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    ExportGeo = vbNullString
    this.lastFailure = vbNullString

    On Error GoTo RecordFailure

    atStep = "reading the Geo sheet"
    Set geo = GeoObject()

    atStep = "adding the output workbook"
    Set outwb = CreateOutputWorkbook()

    atStep = "reading the geobase name"
    geobaseName = geo.GeoBaseName()
    If LenB(geobaseName) = 0 Then
        geobaseName = "OUTBREAK-TOOLS-GEOBASE-UNKNOWN"
    End If

    geobaseName = geobaseName & "_export" & Format$(Now, "yyyymmdd")
    If onlyHistoric Then geobaseName = geobaseName & "_historic"

    filePath = BuildFilePath(folderPath, geobaseName, "xlsx")

    ' Use LLGeo.ExportToFile to write the geobase data
    atStep = "writing the geobase tables"
    geo.ExportToFile outwb, onlyHistoric

    'A geobase export that WORKED can still have left the headers alone, and
    'LLGeo says so in its own account. That class is hand-exported into the
    'designer and reaches no log on purpose, so the log entry is written here,
    'where one is already held.
    ReportGeoWarning geo

    atStep = "saving " & filePath
    SaveWorkbook outwb, filePath, xlOpenXMLWorkbook, vbNullString
    ExportGeo = filePath
    Exit Function

RecordFailure:
    errNumber = Err.Number
    errSource = Err.Source
    errMessage = Err.Description
    RecordExportFailure "geo export", atStep, errNumber, errSource, errMessage

    'The same loss happens one class further in, so LLGeo keeps its own account
    'and it is the one that names what refused.
    If Not geo Is Nothing Then
        If LenB(geo.LastFailure) > 0 Then _
            this.lastFailure = this.lastFailure & " | " & geo.LastFailure
    End If

    Err.Raise errNumber, errSource, errMessage
End Function

'@label:close-all
'@sub-title Close the workbooks this instance still holds open
'@details
'Closes any output workbook left open after an export error, preventing
'orphaned minimized windows, and closes the source workbook when this instance
'opened it through CreateFromFile. Nothing is saved on either close. Safe to
'call even when no workbook is open.
'@export
Public Sub CloseAll()
    On Error Resume Next
    If Not this.activeWkb Is Nothing Then
        this.activeWkb.Close savechanges:=False
        Set this.activeWkb = Nothing
    End If
    If this.openedSource And Not this.sourceWkb Is Nothing Then
        this.sourceWkb.Close savechanges:=False
        Set this.sourceWkb = Nothing
    End If
    On Error GoTo 0
End Sub


'@sub-title Get the sheet type tag from worksheet-level HiddenNames.
'@param sh Worksheet. The worksheet to query.
'@return String. The sheet type tag (HList, VList, etc.).
Private Function SheetTagOf(ByVal sh As Worksheet, _
                            Optional ByVal nameKey As String = "sheet_type") As String
    Dim shHn As HiddenNames
    Set shHn = HiddenNames.Create(sh)
    SheetTagOf = shHn.ValueAsString(nameKey)
End Function

'@section Private - Source Object Accessors
'===============================================================================

'@label:source-dictionary
'@sub-title Get a dictionary from the source workbook
'@param nbExp Long. Number of export columns to configure.
'@return LLdictionary. Dictionary instance from the Dictionary sheet.
Private Function SourceDictionary(ByVal nbExp As Long) As LLdictionary
    Set SourceDictionary = LLdictionary.Create( _
        this.sourceWkb.Worksheets(DICTIONARYSHEET), 1, 1, nbExp)
End Function

'@label:export-object
'@sub-title Get the export configuration from the source workbook
'@return LLExport. Export configuration from the Exports sheet.
Private Function ExportObject() As LLExport
    Set ExportObject = LLExport.Create(this.sourceWkb.Worksheets(EXPORTSHEET))
End Function

'@label:passwords-object
'@sub-title Get passwords from the source workbook
'@return Passwords. Passwords instance from the __pass sheet.
'@details
'Built once and held, and answering Nothing rather than raising when the file
'will not take one.
'
'A LINELIST OLDER THAN THE PROTECTION MATRIX
'-------------------------------------------------------------------------------
'Passwords.Create validates the sheet and demands T_ProtectedSheets, a table
'this version added. A linelist generated before it carries T_Keys, the five
'named ranges and no protection matrix, so the factory refuses a password sheet
'that is complete for everything an export wants from it. The class itself is
'not loosened -- it is baked into the workbook rather than imported, so a change
'there reaches nothing until every workbook is re-pasted by hand. The tolerance
'lives here instead, in a class that imports.
'
'What the absence costs is small and bounded. An export-all file name is
'composed without a passwords object at all; a numbered export loses its version
'suffix and cannot be given the private key. Both are guarded where they read.
'@return Passwords. The held handler, or Nothing when the file carries none.
Private Function PasswordsObject() As Passwords

    Dim refusal As String

    If this.triedPass Then
        Set PasswordsObject = this.passObj
        Exit Function
    End If

    this.triedPass = True

    On Error Resume Next
        Set this.passObj = Passwords.Create(this.sourceWkb.Worksheets(PASSWORDSHEET))
        refusal = Err.Description
    On Error GoTo 0

    If this.passObj Is Nothing Then ReportPasswordsRefusal refusal

    Set PasswordsObject = this.passObj
End Function

'@label:report-geo-warning
'@sub-title Log what the geobase export got past rather than stopped on
'@details
'Read after an export that succeeded. LLGeo fills its LastFailure on a refusal
'it carried on from -- the headers it could not put back -- and leaves it empty
'otherwise, so a value here is a warning and never the reason for a failure.
'On a failure the RecordFailure handler already folds the same account into
'this class''s own, so nothing is said twice.
'@param geo LLGeo. The manager that just wrote the geobase.
Private Sub ReportGeoWarning(ByVal geo As LLGeo)

    Dim detail As String
    Dim logStore As LLLog

    If geo Is Nothing Then Exit Sub

    detail = geo.LastFailure
    If LenB(detail) = 0 Then Exit Sub

    Set logStore = UserLog()
    If logStore Is Nothing Then Exit Sub

    'A log that will not take the entry is not worth losing the export over.
    On Error Resume Next
    logStore.LogWarning "export-geo", detail, "LLGeo.Translate"
    On Error GoTo 0
End Sub

'@label:report-passwords-refusal
'@sub-title Say in the workbook''s log that this file carries no usable passwords
'@details
'Written as a warning rather than a failure: the export goes on to write its
'file, and the user needs to know why the name lost its version suffix and why
'no export of this file can be password protected.
'@param refusal String. What Passwords.Create said when it refused.
Private Sub ReportPasswordsRefusal(ByVal refusal As String)

    Dim detail As String
    Dim logStore As LLLog

    detail = "This file carries no usable password sheet: " & refusal & _
             ". The export runs without one, so its file name has no version " & _
             "suffix and no export of this file can be password protected."

    Set logStore = UserLog()
    If logStore Is Nothing Then Exit Sub

    'A log that will not take the entry is not worth losing the export over.
    On Error Resume Next
    logStore.LogWarning "export-passwords", detail, "LLExporter.PasswordsObject"
    On Error GoTo 0
End Sub

'@label:user-log
'@sub-title The log of the workbook being exported, built once
'@details
'Built on first use and held, the way FilteredData and LLGeo hold theirs. A
'workbook that will not take a log answers Nothing, and the Nothing is the
'state rather than a raise.
'@return LLLog. The held log, or Nothing.
Private Function UserLog() As LLLog

    If Not this.triedLog Then
        this.triedLog = True
        On Error Resume Next
        Set this.userLog = LLLog.Create(this.sourceWkb)
        On Error GoTo 0
    End If

    Set UserLog = this.userLog
End Function

'@label:geo-object
'@sub-title Get the geo object from the source workbook
'@return LLGeo. Geo instance from the Geo sheet.
'@label:record-export-failure
'@sub-title Keep in words what an export refused on, before the raise leaves
'@details
'An error loses its words on the way out of this class: the caller reads the
'name of the method it called and nothing of what actually refused. Every
'export therefore records its own account here, and the caller reads it back
'through LastFailure.
'
'ExportGeo carried this alone for a while, which is why a failing migration
'answered an empty LastFailure and the form had nothing to log.
'@param what String. Which export this is, for the opening words.
'@param atStep String. The step the export was on.
'@param errNumber Long. The error number.
'@param errSource String. The error source.
'@param errMessage String. The error description.
Private Sub RecordExportFailure(ByVal what As String, ByVal atStep As String, _
                                ByVal errNumber As Long, _
                                ByVal errSource As String, _
                                ByVal errMessage As String)

    this.lastFailure = what & " failed " & atStep & " - error " & _
                       errNumber & " from " & errSource & ": " & errMessage
End Sub

Private Function GeoObject() As LLGeo
    'The checks are off here on purpose. A linelist older than the hidden name
    'store keeps its geo metadata in plain named ranges, and HiddenNames tracks
    'a name only when it is hidden, so those four names read as absent and the
    'factory refuses a geobase whose tables are all present. An export reads the
    'tables and writes them out; what the file cannot answer fails at the read
    'that wants it, and that read names the missing piece.
    Set GeoObject = LLGeo.Create(this.sourceWkb.Worksheets(GEOSHEET), _
                                 runChecks:=False)
End Function


'@section Private - Output Workbook Management
'===============================================================================

'@label:create-output-workbook
'@sub-title Create a new minimized output workbook
'@return Workbook. The newly created workbook.
Private Function CreateOutputWorkbook() As Workbook
    Set this.activeWkb = Workbooks.Add
    ActiveWindow.WindowState = xlMinimized
    Set CreateOutputWorkbook = this.activeWkb
End Function

'@label:save-workbook
'@sub-title Save and close the output workbook
'@details
'Saves the workbook to the specified path with the given format and
'optional password, then closes it without saving again.
'@param outwb Workbook. The workbook to save.
'@param filePath String. Full path including file name and extension.
'@param fileFormat Long. Excel file format constant (xlOpenXMLWorkbook or xlExcel12).
'@param password String. Password to apply, or vbNullString for none.
Private Sub SaveWorkbook(ByVal outwb As Workbook, ByVal filePath As String, _
                         ByVal fileFormat As Long, ByVal password As String)

    If LenB(password) > 0 Then
        outwb.SaveAs fileName:=filePath, _
                     fileFormat:=fileFormat, _
                     CreateBackup:=False, _
                     password:=password, _
                     ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    Else
        outwb.SaveAs fileName:=filePath, _
                     fileFormat:=fileFormat, _
                     CreateBackup:=False, _
                     ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    End If

    outwb.Close savechanges:=False
    Set this.activeWkb = Nothing
End Sub

'@label:build-file-path
'@sub-title Build a full file path from folder, name, and extension
'@param folderPath String. Target folder path.
'@param fileName String. File name without extension.
'@param fileExt String. File extension without dot.
'@return String. The assembled file path.
Private Function BuildFilePath(ByVal folderPath As String, ByVal fileName As String, _
                               ByVal fileExt As String) As String
    Dim cleanPath As String

    cleanPath = folderPath
    ' Remove trailing separator if present
    If InStrRev(cleanPath, Application.PathSeparator, Len(cleanPath)) = Len(cleanPath) Then
        cleanPath = Left$(cleanPath, Len(cleanPath) - 1)
    End If

    BuildFilePath = cleanPath & Application.PathSeparator & fileName & "." & fileExt
End Function


'@section Private - Dictionary Preparation
'===============================================================================

'@label:excluded-variables
'@sub-title The variables a custom export leaves out
'@details
'Reads three cells of the export row and answers the names of the variables the
'export drops: the geo variables of every admin level "admin levels" leaves out,
'the geo p-code and geo concat variables when "include p-codes" is off, and
'every variable flagged as a personal identifier when "include personal
'identifiers" is off.
'
'ONE LEVEL OWNS THREE ROWS
'-------------------------------------------------------------------------------
'LLdictionary.AppendGeoLines writes a visible name row, a hidden p-code row and
'a hidden concat row for each of the four levels, and a level left out drops all
'three. The health facility rows are left alone: `hf` and `geo_pcode_hf` belong
'to no admin level, and an export keeps them whatever the cell asks for.
'
'WHY THIS READS AND NEVER WRITES
'-------------------------------------------------------------------------------
'The exclusion used to be done by clearing the export column of the linelist's
'own Dictionary worksheet, so one export with p-codes off took those variables
'out of that export for good and turning the flag back on could not bring them
'back. The dictionary is the file the whole linelist is generated from and it is
'what a migration ships, so the loss travelled to the next linelist. The names
'answered here are dropped from the copies the export builds, and the source
'sheet keeps what the setup put in it.
'@param dict LLdictionary. The source dictionary.
'@param expObj LLExport. The export configuration.
'@param exportNumber Long. Which export definition to read.
'@return BetterArray. The variable names to drop, empty when the export keeps
'everything.
Private Function ExcludedVariables(ByVal dict As LLdictionary, ByVal expObj As LLExport, _
                                   ByVal exportNumber As Long) As BetterArray

    Dim excluded As BetterArray
    Dim pcodesTagList As BetterArray
    Dim varNames As Variant
    Dim mainLabels As Variant
    Dim identifiers As Variant
    Dim controlValues As Variant
    Dim keptLevels() As Boolean
    Dim dropPcodes As Boolean
    Dim dropIdentifiers As Boolean
    Dim dropLevels As Boolean
    Dim rowLevel As Long
    Dim counter As Long

    Set excluded = New BetterArray
    Set ExcludedVariables = excluded

    dropPcodes = (expObj.ColumnValue(exportNumber, "include p-codes") <> "yes")
    dropIdentifiers = (expObj.ColumnValue(exportNumber, _
                                          "include personal identifiers") <> "yes")
    keptLevels = KeptAdminLevels(expObj.ColumnValue(exportNumber, "admin levels"), _
                                 dropLevels)

    If Not dropPcodes And Not dropIdentifiers And Not dropLevels Then Exit Function

    varNames = ColumnValues(dict.DataRange("variable name"))
    If dropPcodes Or dropLevels Then _
        mainLabels = ColumnValues(dict.DataRange("main label"))

    If dropPcodes Then
        Set pcodesTagList = New BetterArray
        pcodesTagList.Push PCODES_GEO_PCODE_ADM1, PCODES_GEO_PCODE_ADM2, _
                           PCODES_GEO_PCODE_ADM3, PCODES_GEO_PCODE_ADM4, _
                           PCODES_GEO_CONCAT_ADM1, PCODES_GEO_CONCAT_ADM2, _
                           PCODES_GEO_CONCAT_ADM3, PCODES_GEO_CONCAT_ADM4

        For counter = LBound(varNames) To UBound(varNames)
            If pcodesTagList.Includes(CStr(mainLabels(counter))) Then
                excluded.Push CStr(varNames(counter))
            End If
        Next counter
    End If

    'A NAME THIS PASS FINDS MAY ALREADY BE ON THE LIST
    '---------------------------------------------------------------------------
    'An export that turns p-codes off and asks for two levels drops the p-code
    'and the concat of the levels it left out twice over. The list is what
    'ExcludedVariablesFor shows a user, so each name is pushed once.
    If dropLevels Then
        controlValues = ColumnValues(dict.DataRange("control"))
        For counter = LBound(varNames) To UBound(varNames)
            rowLevel = GeoLevelOf(CStr(controlValues(counter)), _
                                  CStr(mainLabels(counter)))
            If rowLevel > 0 Then
                If Not keptLevels(rowLevel) Then
                    If Not excluded.Includes(CStr(varNames(counter))) Then _
                        excluded.Push CStr(varNames(counter))
                End If
            End If
        Next counter
    End If

    If dropIdentifiers Then
        identifiers = ColumnValues(dict.DataRange("personal identifier"))
        For counter = LBound(varNames) To UBound(varNames)
            If CStr(identifiers(counter)) = "yes" Then
                excluded.Push CStr(varNames(counter))
            End If
        Next counter
    End If
End Function

'@label:kept-admin-levels
'@sub-title The admin levels one export row asks for
'@details
'The "admin levels" cell holds the levels joined with a plus, written
'`admin1+admin2`. An empty cell asks for every level.
'
'THE CELL IS HONOURED ONLY WHEN THE WHOLE OF IT READS
'-------------------------------------------------------------------------------
'A cell holding one chunk this cannot read answers every level, the way an empty
'cell does. An export that silently drops geo columns over a typo is worse than
'one that keeps them all, and SetupErrors.CheckExports reports that same cell on
'the setup side, before a linelist is ever built from it.
'@param cellText String. What the Exports row holds.
'@param someDropped Boolean. Written True when a level is left out.
'@return Boolean(). One entry per level, 1 to 4, True for a level to export.
Private Function KeptAdminLevels(ByVal cellText As String, _
                                 ByRef someDropped As Boolean) As Boolean()

    Dim kept() As Boolean
    Dim chunks As Variant
    Dim levelNumber As Long
    Dim counter As Long
    Dim readable As Boolean

    ReDim kept(1 To ADMIN_LEVEL_COUNT)
    someDropped = False
    readable = (LenB(Trim$(cellText)) > 0)

    If readable Then
        chunks = Split(cellText, ADMIN_LEVEL_SEPARATOR)
        For counter = LBound(chunks) To UBound(chunks)
            levelNumber = AdminLevelOfChunk(CStr(chunks(counter)))
            If levelNumber > 0 Then
                kept(levelNumber) = True
            Else
                readable = False
                Exit For
            End If
        Next counter
    End If

    For counter = 1 To ADMIN_LEVEL_COUNT
        If readable Then
            If Not kept(counter) Then someDropped = True
        Else
            kept(counter) = True
        End If
    Next counter

    KeptAdminLevels = kept
End Function

'@label:admin-level-of-chunk
'@sub-title The level one chunk of the "admin levels" cell holds
'@details
'The chunk is trimmed and read without regard to case, so `Admin2` and ` admin2`
'both answer 2. Anything else answers 0, and the caller reads that as a cell it
'cannot honour. SetupErrors.IsAdminLevelsCell holds this same rule.
'@param chunk String. One piece of the cell, between two plus signs.
'@return Long. The level, 1 to 4, or 0 when the chunk holds something else.
Private Function AdminLevelOfChunk(ByVal chunk As String) As Long

    Dim chunkText As String
    Dim tail As String

    chunkText = Trim$(chunk)
    If Len(chunkText) <> Len(ADMIN_LEVEL_PREFIX) + 1 Then Exit Function
    If StrComp(Left$(chunkText, Len(ADMIN_LEVEL_PREFIX)), _
               ADMIN_LEVEL_PREFIX, vbTextCompare) <> 0 Then Exit Function

    tail = Right$(chunkText, 1)
    If tail < "1" Or tail > CStr(ADMIN_LEVEL_COUNT) Then Exit Function

    AdminLevelOfChunk = CLng(tail)
End Function

'@label:geo-level-of
'@sub-title The admin level one dictionary row belongs to
'@details
'LLdictionary.AppendGeoLines writes three rows per level and this reads all
'three back. The visible name row carries the control `geo1` to `geo4`. The two
'hidden rows carry the control `formula`, and their level sits in the main
'label, written `geo_pcode_adm2` or `geo_concat_adm2`.
'A row belonging to no admin level answers 0, and the health facility rows are
'among them: `hf` and `geo_pcode_hf` name no level.
'@param controlValue String. The `control` cell of the row.
'@param mainLabel String. The `main label` cell of the row.
'@return Long. The level, 1 to 4, or 0 when the row names none.
Private Function GeoLevelOf(ByVal controlValue As String, _
                            ByVal mainLabel As String) As Long

    Dim controlText As String
    Dim labelText As String
    Dim counter As Long

    controlText = LCase$(Trim$(controlValue))
    labelText = LCase$(Trim$(mainLabel))

    For counter = 1 To ADMIN_LEVEL_COUNT
        If controlText = "geo" & counter Then
            GeoLevelOf = counter
            Exit Function
        End If
        If labelText = "geo_pcode_adm" & counter Or _
           labelText = "geo_concat_adm" & counter Then
            GeoLevelOf = counter
            Exit Function
        End If
    Next counter
End Function

'@label:column-values
'@sub-title Read a one-column range into a 1-based array
'@details
'A range of one cell answers a scalar rather than an array, so that case is
'built by hand. Everything else is one crossing for the whole column.
'@param rng Range. A range one column wide.
'@return Variant. A 1-based array of the cell values.
Private Function ColumnValues(ByVal rng As Range) As Variant
    Dim block As Variant
    Dim values() As Variant
    Dim rowCount As Long
    Dim counter As Long

    rowCount = rng.Rows.count
    ReDim values(1 To rowCount)

    If rowCount = 1 Then
        values(1) = rng.Cells(1, 1).Value
    Else
        block = rng.Value
        For counter = 1 To rowCount
            values(counter) = block(counter, 1)
        Next counter
    End If

    ColumnValues = values
End Function

'@label:without-excluded-rows
'@sub-title Drop the excluded variables from a dictionary block
'@details
'The block carries its headers on the first row, so the variable name column is
'found by its header and every row below is kept or dropped on the name it
'holds. The block is answered unchanged when the export excludes nothing or
'when the block carries no variable name column.
'@param dictData BetterArray. A dictionary block, headers included.
'@param excluded BetterArray. The variable names to drop.
'@return BetterArray. The block with the excluded rows removed.
Private Function WithoutExcludedRows(ByVal dictData As BetterArray, _
                                     ByVal excluded As BetterArray) As BetterArray

    Dim block As Variant
    Dim kept() As Variant
    Dim trimmed As BetterArray
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim nameColumn As Long
    Dim keptRow As Long
    Dim rowCounter As Long
    Dim colCounter As Long

    Set WithoutExcludedRows = dictData
    If excluded.Length = 0 Then Exit Function

    'A block of one row comes back one-dimensional and holds the headers alone,
    'so there is nothing to drop from it.
    If dictData.ArrayType <> ArrayTypes.BA_MULTIDIMENSION Then Exit Function

    block = dictData.Items
    firstRow = LBound(block, 1)
    lastRow = UBound(block, 1)
    firstCol = LBound(block, 2)
    lastCol = UBound(block, 2)

    nameColumn = firstCol - 1
    For colCounter = firstCol To lastCol
        If LCase$(Trim$(CStr(block(firstRow, colCounter)))) = "variable name" Then
            nameColumn = colCounter
            Exit For
        End If
    Next colCounter
    If nameColumn < firstCol Then Exit Function

    ReDim kept(firstRow To lastRow, firstCol To lastCol)
    keptRow = firstRow

    For rowCounter = firstRow To lastRow
        If rowCounter = firstRow Or _
           Not excluded.Includes(CStr(block(rowCounter, nameColumn))) Then
            For colCounter = firstCol To lastCol
                kept(keptRow, colCounter) = block(rowCounter, colCounter)
            Next colCounter
            keptRow = keptRow + 1
        End If
    Next rowCounter

    Set trimmed = New BetterArray
    trimmed.LowerBound = firstRow
    trimmed.Items = TopRowsOf(kept, firstRow, keptRow - 1, firstCol, lastCol)
    Set WithoutExcludedRows = trimmed
End Function

'@label:top-rows-of
'@sub-title Copy the first rows of a two-dimensional block
'@details
'ReDim Preserve reaches the last dimension alone, so a block that lost rows is
'copied into a block of the right height instead.
'@param block Variant. The source block.
'@param firstRow Long. First row to copy.
'@param lastRow Long. Last row to copy.
'@param firstCol Long. First column of the block.
'@param lastCol Long. Last column of the block.
'@return Variant. A block holding the rows asked for.
Private Function TopRowsOf(ByVal block As Variant, ByVal firstRow As Long, _
                           ByVal lastRow As Long, ByVal firstCol As Long, _
                           ByVal lastCol As Long) As Variant
    Dim copied() As Variant
    Dim rowCounter As Long
    Dim colCounter As Long

    ReDim copied(firstRow To lastRow, firstCol To lastCol)

    For rowCounter = firstRow To lastRow
        For colCounter = firstCol To lastCol
            copied(rowCounter, colCounter) = block(rowCounter, colCounter)
        Next colCounter
    Next rowCounter

    TopRowsOf = copied
End Function

'@label:prepare-temp-dictionary
'@sub-title Write dictionary data to the temp sheet for export processing
'@details
'When fullExport is True, copies the entire dictionary. When False,
'filters on the export column name to include only rows assigned to that
'export. The variables the export excludes are dropped from the block in
'memory. Writes the result to the __temp sheet and adds a blue font tag
'to mark the dictionary as prepared.
'@param dict LLdictionary. The source dictionary.
'@param expName String. The export column name to filter on, or vbNullString for full.
'@param fullExport Boolean. When True, copies the entire dictionary unfiltered.
'@param excluded BetterArray. Variable names to drop from the copy.
Private Sub PrepareTempDictionary(ByVal dict As LLdictionary, _
                                  ByVal expName As String, _
                                  ByVal fullExport As Boolean, _
                                  ByVal excluded As BetterArray)

    Dim dictData As BetterArray
    Dim tempsh As Worksheet

    Set tempsh = InternalSheetOf(this.sourceWkb, TEMPSHEET, TEMPSHEET_OLD)

    If fullExport Then
        Set dictData = New BetterArray
        dictData.FromExcelRange dict.DataRange(includeHeaders:=True)
    Else
        Set dictData = dict.Data.FilterData( _
            varName:=expName, criteriaName:="<>", _
            returnedColumnName:="__all__", includeHeaders:=True, _
            matchCase:=False)
    End If

    Set dictData = WithoutExcludedRows(dictData, excluded)

    tempsh.Cells.Clear
    dictData.ToExcelRange tempsh.Cells(1, 1)

    ' Add preparation tag (blue font on next row signals prepared dictionary)
    Dim expDictRng As Range
    Set expDictRng = LLdictionary.Create(tempsh, 1, 1, _
        ExportObject().NumberOfExports()).DataRange(includeHeaders:=True)
    expDictRng.Cells(expDictRng.Rows.count + 1, 1).Font.Color = vbBlue
End Sub

'@label:export-dictionary-from-temp
'@sub-title Create an export dictionary from the temp sheet
'@param nbExp Long. Number of export columns.
'@return LLdictionary. Dictionary built from the __temp sheet.
Private Function ExportDictionaryFromTemp(ByVal nbExp As Long) As LLdictionary
    Set ExportDictionaryFromTemp = LLdictionary.Create( _
        InternalSheetOf(this.sourceWkb, TEMPSHEET, TEMPSHEET_OLD), 1, 1, nbExp)
End Function

'@label:sort-export-dictionary
'@sub-title Sort the export dictionary by table name and export column
'@param expDict LLdictionary. The dictionary to sort.
'@param expName String. The export column name to use as secondary sort key.
Private Sub SortExportDictionary(ByVal expDict As LLdictionary, ByVal expName As String)
    Dim rng As Range
    Set rng = expDict.DataRange(includeHeaders:=True)
    rng.Sort _
        key1:=expDict.DataRange("table name", includeHeaders:=True), _
        key2:=expDict.DataRange(expName, includeHeaders:=True), _
        Header:=xlYes
End Sub


'@section Private - Metadata and Sheets
'===============================================================================

'@label:add-metadata-sheets
'@sub-title Add Metadata, Choices, and Translations sheets to the output workbook
'@param outwb Workbook. The output workbook.
'@param dict LLdictionary. The source dictionary.
'@param nbExp Long. Number of export columns.
Private Sub AddMetadataSheets(ByVal outwb As Workbook, ByVal dict As LLdictionary, _
                              ByVal nbExp As Long)

    Dim sheetNames As BetterArray
    Dim csData As DataSheet
    Dim sh As Worksheet
    Dim counter As Long
    Dim sheetName As String

    Set sheetNames = New BetterArray
    sheetNames.Push METADATASHEET, CHOICESSHEET

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)
        Set sh = this.sourceWkb.Worksheets(sheetName)
        Set csData = DataSheet.Create(sh, 1, 1)
        csData.Export toWkb:=outwb, Hide:=xlSheetVisible
    Next counter

    ' Export dictionary translations, from whichever layout the file uses
    AddDictionaryTranslations outwb
End Sub

'@label:add-dictionary-translations
'@sub-title Put the dictionary translations into the output, either layout
'@details
'The current layout keeps them in Tab_Translations on the LinelistTranslation
'sheet, and LLTranslation.ExportDictionary writes them out from there onto a
'sheet called Translations.
'
'THE LAYOUT BEFORE THAT ONE
'-------------------------------------------------------------------------------
'A linelist older than that table keeps its dictionary translations on a
'worksheet of its own, already called Translations, and carries no
'Tab_Translations at all. LLTranslation.Create demands that table and refuses
'the whole sheet without it, so the export used to stop here.
'
'Skipping is not the answer: the Translations sheet of an export file is what
'carries every translated label to whatever reads the file back, and an export
'without one loses all of them silently. So the older layout is read where it
'actually sits and copied straight across, the same way Metadata and Choices
'are copied above.
'@param outwb Workbook. The output workbook being built.
Private Sub AddDictionaryTranslations(ByVal outwb As Workbook)

    Dim translations As LLTranslation
    Dim refusal As String

    On Error Resume Next
        Set translations = LLTranslation.Create( _
            this.sourceWkb.Worksheets(LINELISTTRANSLATIONSHEET))
        refusal = Err.Description
    On Error GoTo 0

    If Not translations Is Nothing Then
        translations.ExportDictionary outwb
        Exit Sub
    End If

    CopyOldTranslationsSheet outwb, refusal
End Sub

'@label:copy-old-translations-sheet
'@sub-title Copy the dictionary translations of an older linelist across
'@details
'The destination name is the one ExportDictionary would have used, so a file
'built this way is shaped like every other export and needs nothing special to
'read it back.
'@param outwb Workbook. The output workbook being built.
'@param refusal String. What LLTranslation.Create said when it refused.
Private Sub CopyOldTranslationsSheet(ByVal outwb As Workbook, _
                                     ByVal refusal As String)

    Dim oldSh As Worksheet
    Dim csData As DataSheet

    On Error Resume Next
    Set oldSh = this.sourceWkb.Worksheets(DICTTRANSLATIONSHEET)
    On Error GoTo 0

    If oldSh Is Nothing Then
        ReportTranslationsWarning "This file carries no dictionary translations " & _
            "in either layout: " & refusal & ", and it has no " & _
            DICTTRANSLATIONSHEET & " worksheet either. The exported file goes " & _
            "out with no translated labels."
        Exit Sub
    End If

    Set csData = DataSheet.Create(oldSh, 1, 1)
    csData.Export toWkb:=outwb, Hide:=xlSheetVisible

    ReportTranslationsWarning "The dictionary translations were read from the " & _
        DICTTRANSLATIONSHEET & " worksheet, the layout used before " & _
        "Tab_Translations: " & refusal & ". The exported file carries them " & _
        "on its " & DICTTRANSLATIONSHEET & " sheet as usual."
End Sub

'@label:report-translations-warning
'@sub-title Log which layout the dictionary translations came from
'@param detail String. What to tell the user.
Private Sub ReportTranslationsWarning(ByVal detail As String)

    Dim logStore As LLLog

    Set logStore = UserLog()
    If logStore Is Nothing Then Exit Sub

    'A log that will not take the entry is not worth losing the export over.
    On Error Resume Next
    logStore.LogWarning "export-translations", detail, "LLExporter.AddDictionaryTranslations"
    On Error GoTo 0
End Sub

'@label:add-dictionary
'@sub-title Add the dictionary to the output workbook
'@details
'Exports the dictionary filtered by exportColumn, drops the rows this export
'excludes from the copy, then marks the output dictionary as prepared with a
'blue font tag.
'
'WHY THE EXCLUDED ROWS ARE DROPPED HERE
'-------------------------------------------------------------------------------
'The export column of the source sheet decides what this copy carries, and that
'column is left as the setup wrote it. So the copy arrives holding the p-codes
'and the personal identifiers the export leaves out, and the rows are taken off
'the copy. The file then lists the variables its data sheets carry, and the
'linelist keeps its own dictionary whole.
'@param outwb Workbook. The output workbook.
'@param dict LLdictionary. The source dictionary.
'@param nbExp Long. Number of export columns.
'@param exportColumn String. The export column to filter on, or "__all__".
'@param excluded BetterArray. Variable names this export leaves out.
Private Sub AddDictionary(ByVal outwb As Workbook, ByVal dict As LLdictionary, _
                          ByVal nbExp As Long, ByVal exportColumn As String, _
                          ByVal excluded As BetterArray)

    Dim outDict As LLdictionary
    Dim outDictRng As Range

    dict.Export toWkb:=outwb, exportType:=exportColumn, _
               addListObject:=False, Hide:=xlSheetVisible

    Set outDict = LLdictionary.Create(outwb.Worksheets(DICTIONARYSHEET), 1, 1, nbExp)
    DropExcludedFromCopy outDict, excluded

    ' Mark the output dictionary as prepared
    Set outDictRng = outDict.DataRange()
    outDictRng.Cells(outDictRng.Rows.count + 1, 1).Font.Color = vbBlue
End Sub

'@label:drop-excluded-from-copy
'@sub-title Take the excluded variables off an exported dictionary
'@details
'The rows are gathered into one range and deleted in a single call, so a
'dictionary with many identifiers costs one crossing rather than one per row.
'@param outDict LLdictionary. The dictionary of the output workbook.
'@param excluded BetterArray. Variable names to remove.
Private Sub DropExcludedFromCopy(ByVal outDict As LLdictionary, _
                                 ByVal excluded As BetterArray)

    Dim namesRng As Range
    Dim deleteRng As Range
    Dim varNames As Variant
    Dim counter As Long

    If excluded.Length = 0 Then Exit Sub

    Set namesRng = outDict.DataRange("variable name")
    varNames = ColumnValues(namesRng)

    For counter = LBound(varNames) To UBound(varNames)
        If excluded.Includes(CStr(varNames(counter))) Then
            If deleteRng Is Nothing Then
                Set deleteRng = namesRng.Cells(counter, 1)
            Else
                Set deleteRng = Application.Union(deleteRng, namesRng.Cells(counter, 1))
            End If
        End If
    Next counter

    If Not deleteRng Is Nothing Then deleteRng.EntireRow.Delete
End Sub

'@label:add-custom-dropdowns
'@sub-title Add custom dropdowns to the Choices sheet in the output workbook
'@details
'Walks the dropdown registry of the source workbook and writes the values of
'each custom dropdown onto the output Choices sheet. Silently exits when no
'custom dropdown sheet exists.
'
'The walk goes through DropdownLists.AllDropdowns, so the key carries the
'dropdown name the registry holds. The header cell of the ListObject was read
'before, and that cell holds the name with every space turned into an
'underscore, so a dropdown called `contact type` went into the file under a name
'the import side could not update. ChoiceKeys builds the key both ends use.
'@param outwb Workbook. The output workbook.
Private Sub AddCustomDropdowns(ByVal outwb As Workbook)
    Dim wkbNames As HiddenNames
    Dim choi As LLChoices
    Dim drop As DropdownLists
    Dim keys As ChoiceKeys
    Dim dropNames As BetterArray
    Dim loValues As BetterArray
    Dim sh As Worksheet
    Dim listName As String
    Dim counter As Long

    Set wkbNames = HiddenNames.Create(this.sourceWkb)

    On Error Resume Next
    Set sh = this.sourceWkb.Worksheets(wkbNames.ValueAsString("RNG_CustomDrop"))
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    Set choi = LLChoices.Create(outwb.Worksheets(CHOICESSHEET), 1, 1)
    Set drop = DropdownLists.Create(sh, vbNullString)
    Set keys = ChoiceKeys.Create()
    Set dropNames = drop.AllDropdowns

    For counter = dropNames.LowerBound To dropNames.UpperBound
        listName = CStr(dropNames.Item(counter))
        Set loValues = drop.Values(listName)
        If loValues.Length > 0 Then
            choi.AddChoice keys.CustomChoiceName(listName), loValues
        End If
    Next counter
End Sub

'@label:add-metadata-tags
'@sub-title Add metadata tags to the Metadata sheet in the output workbook
'@details
'Appends export_type, export_date, editable_labels, show_hide, and
'epiweekstart tags to the Metadata sheet variable/value table. These tags
'control import behaviour when the exported file is re-imported.
'@param outwb Workbook. The output workbook.
'@param scope Byte. ExportScope value (ExportScopeCustom or ExportScopeMigration).
'@param exportNumber Long. The export number for custom exports.
'@param includeShowHide Boolean. When True, tags show_hide as "update on import".
'@param keepLabels Boolean. When True, tags editable_labels as "update on import".
'@param useFilters Boolean. When True, the data came from the filtered companions.
Private Sub AddMetadataTags(ByVal outwb As Workbook, ByVal scope As Byte, _
                            ByVal exportNumber As Long, _
                            ByVal includeShowHide As Boolean, _
                            ByVal keepLabels As Boolean, _
                            ByVal useFilters As Boolean)

    Dim csData As DataSheet
    Dim varRng As Range
    Dim expName As String
    Dim rowOffset As Long

    Set csData = DataSheet.Create(outwb.Worksheets(METADATASHEET), 1, 1)
    Set varRng = csData.DataRange("variable")
    rowOffset = varRng.Rows.count

    ' Export type tag
    If scope = ExportScopeMigration Then
        expName = "export for migrations"
    Else
        expName = "custom export: export " & CStr(exportNumber)
    End If
    varRng.Cells(rowOffset + 1, 1).Value = "export_type"
    varRng.Cells(rowOffset + 1, 2).Value = expName

    ' Export date
    varRng.Cells(rowOffset + 2, 1).Value = "export_date"
    varRng.Cells(rowOffset + 2, 2).Value = Format$(Now, "yyyy-mm-dd Hh:Nn")

    ' Editable labels tag (default: no update on import)
    varRng.Cells(rowOffset + 3, 1).Value = "editable_labels"
    varRng.Cells(rowOffset + 3, 2).Value = IIf(keepLabels, "update on import", "no update on import")

    ' Show/hide tag (default: no update on import)
    varRng.Cells(rowOffset + 4, 1).Value = "show_hide"
    varRng.Cells(rowOffset + 4, 2).Value = IIf(includeShowHide, "update on import", "no update on import")

    ' Epiweek start value. A local of that name would shadow the property that
    ' answers it, so the property is called straight into the cell.
    varRng.Cells(rowOffset + 5, 1).Value = "epiweekstart"
    varRng.Cells(rowOffset + 5, 2).Value = EpiWeekStart()

    ' The options this export was written with, so the import can read the file
    ' knowing what is in it.
    varRng.Cells(rowOffset + 6, 1).Value = "export_options"
    varRng.Cells(rowOffset + 6, 2).Value = ExportOptionsTag(scope, exportNumber, useFilters)

    ' The workbook the file came out of. The four generation rows the linelist
    ' carries - linelist_creation_date, linelist_creation_name, used_setup_name
    ' and used_designer_name - travel with the copied Metadata sheet, and they
    ' name the workbook as it was built. This names it as it is today, which is
    ' what a person holding the file wants when the linelist has been renamed.
    varRng.Cells(rowOffset + 7, 1).Value = "source_linelist"
    varRng.Cells(rowOffset + 7, 2).Value = this.sourceWkb.Name
End Sub

'@label:export-options-tag
'@sub-title The options of one export, written as one line the import reads back
'@details
'The Exports worksheet stays in the linelist, so a file has no way of saying how
'it was written unless the export writes it down. Every option is a
'`name: value` pair and the pairs are joined with a pipe, so the import side
'splits on the pipe and then on the colon.
'
'The two pairs the import cannot work without are `variable names` and
'`header row`. A file whose header row carries labels alone holds no variable
'names anywhere, so it can never be read back; a file carrying both has its
'variable names on row 2 and reads back once the import knows to look there.
'@param scope Byte. ExportScope value.
'@param exportNumber Long. The export number, ignored for a migration.
'@param useFilters Boolean. Whether the data came from the filtered companions.
'@return String. The options of this export on one line.
Private Function ExportOptionsTag(ByVal scope As Byte, ByVal exportNumber As Long, _
                                  ByVal useFilters As Boolean) As String

    Dim expObj As LLExport
    Dim options As BetterArray
    Dim headerFormat As String

    Set options = New BetterArray
    options.LowerBound = 1

    If scope = ExportScopeMigration Then
        options.Push OnePair("header format", HEADER_NAMES)
        options.Push OnePair("variable names", "yes")
        options.Push OnePair("variable labels", "no")
        options.Push OnePair("header row", "1")
        options.Push OnePair("filtered data", "no")
        options.Push OnePair("admin levels", ADMIN_LEVELS_ALL)
        ExportOptionsTag = Join(options.Items, OPTION_SEPARATOR)
        Exit Function
    End If

    Set expObj = ExportObject()
    headerFormat = expObj.ColumnValue(exportNumber, "header format")
    If LenB(headerFormat) = 0 Then headerFormat = HEADER_NAMES

    options.Push OnePair("header format", headerFormat)
    options.Push OnePair("variable names", _
                         IIf(headerFormat = HEADER_LABELS, "no", "yes"))
    options.Push OnePair("variable labels", _
                         IIf(headerFormat = HEADER_NAMES, "no", "yes"))
    options.Push OnePair("header row", _
                         IIf(headerFormat = HEADER_NAMES_LABELS, "2", "1"))
    options.Push OnePair("filtered data", IIf(useFilters, "yes", "no"))
    options.Push OnePair("file format", expObj.ColumnValue(exportNumber, "file format"))
    options.Push OnePair("password", expObj.ColumnValue(exportNumber, "password"))
    options.Push OnePair("include personal identifiers", _
                         expObj.ColumnValue(exportNumber, "include personal identifiers"))
    options.Push OnePair("include p-codes", _
                         expObj.ColumnValue(exportNumber, "include p-codes"))
    options.Push OnePair("admin levels", _
                         expObj.ColumnValue(exportNumber, "admin levels"), _
                         whenEmpty:=ADMIN_LEVELS_ALL)
    options.Push OnePair("export metadata sheets", _
                         expObj.ColumnValue(exportNumber, "export metadata sheets"))
    options.Push OnePair("export analyses sheets", _
                         expObj.ColumnValue(exportNumber, "export analyses sheets"))

    ExportOptionsTag = Join(options.Items, OPTION_SEPARATOR)
End Function

'@label:one-pair
'@sub-title One `name: value` pair of the export options line
'@details
'An option the Exports row left empty is written as `no`, so the reader on the
'other side never has to tell an absent option from a false one. An option whose
'empty cell means something else says so: "admin levels" left empty asks for
'every level, and `no` would read as none of them.
'@param optionName String. The option name.
'@param optionValue String. What the Exports row holds for it.
'@param whenEmpty Optional String. What an empty cell is written as. Defaults to `no`.
'@return String. The pair.
Private Function OnePair(ByVal optionName As String, ByVal optionValue As String, _
                         Optional ByVal whenEmpty As String = "no") As String
    Dim writtenValue As String

    writtenValue = optionValue
    If LenB(Trim$(writtenValue)) = 0 Then writtenValue = whenEmpty

    OnePair = optionName & OPTION_PAIR_SEPARATOR & writtenValue
End Function

'@label:add-show-hide
'@sub-title Export the show/hide choices with their sizes to the output workbook
'@details
'Builds a __show_hide sheet in the output workbook and asks ShowHideStore to
'provision its table. Every data sheet contributes each layer it has: HList and
'CRF for a data entry sheet, Printed for its companion, VList for an info sheet.
'The choices the user made are read out of the source workbook first, so the
'file carries what the user sees rather than what the dictionary authored.
'
'The CRF layer used to be left out of both ends of the migration, so a CRF
'sheet lost its visibility and its sizes on every round trip.
'@param outwb Workbook. The output workbook.
Private Sub AddShowHide(ByVal outwb As Workbook)

    Dim outsh As Worksheet
    Dim dict As LLdictionary
    Dim sheetNames As BetterArray
    Dim sheetName As String
    Dim sheetType As String
    Dim sh As Worksheet
    Dim printSh As Worksheet
    Dim crfSh As Worksheet
    Dim outStore As ShowHideStore
    Dim srcSh As Worksheet
    Dim srcStore As ShowHideStore
    Dim counter As Long
    Dim nbExp As Long

    nbExp = ExportObject().NumberOfExports()
    Set dict = SourceDictionary(nbExp)

    'The first positional argument of Worksheets.Add is Before, so this used to
    'put the sheet in front of the last one while reading as "after".
    outwb.Worksheets.Add(After:=outwb.Worksheets(outwb.Worksheets.count)).Name = SHOWHIDESHEET
    Set outsh = outwb.Worksheets(SHOWHIDESHEET)
    Set outStore = ShowHideStore.CreateOnSheet(outsh)

    Set sheetNames = dict.UniqueValues("sheet name")

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)

        Set sh = Nothing
        On Error Resume Next
        Set sh = this.sourceWkb.Worksheets(sheetName)
        On Error GoTo 0
        If sh Is Nothing Then GoTo NextSheet

        sheetType = SheetTagOf(sh)

        If sheetType = "HList" Then
            SaveOneShowHideLayer outStore, dict, ShowHideLayerHList, sheetName, sh

            Set printSh = CompanionSheet(PRINTPREFIX & sheetName)
            If Not printSh Is Nothing Then
                SaveOneShowHideLayer outStore, dict, ShowHideLayerPrinted, sheetName, printSh
            End If

            Set crfSh = CompanionSheet(CRFPREFIX & sheetName)
            If Not crfSh Is Nothing Then
                SaveOneShowHideLayer outStore, dict, ShowHideLayerCRF, sheetName, crfSh
            End If

        ElseIf sheetType = "VList" Then
            SaveOneShowHideLayer outStore, dict, ShowHideLayerVList, sheetName, sh
        End If

NextSheet:
    Next counter

    'The saved layouts travel with the current state. The source store is read
    'as it stands; a linelist with no store sheet has nothing to carry, and the
    'export table starts empty of layouts, so nothing is ever skipped here.
    Set srcSh = InternalSheetOf(this.sourceWkb, SHOWHIDESHEET, SHOWHIDESHEET_OLD)
    Set srcStore = ShowHideStore.CreateForRead(srcSh)
    outStore.MergeLayoutsFrom srcStore

    On Error Resume Next
    outsh.Cells.EntireColumn.AutoFit
    On Error GoTo 0
End Sub

'@label:save-one-show-hide-layer
'@sub-title Write one layer of one sheet into the export store
'@details
'Reads the entry list from the dictionary, folds in whatever the user chose in
'the running workbook, then saves it with the sheet so the sizes and the header
'directions travel too.
'@param outStore ShowHideStore. The store on the output workbook.
'@param dict LLdictionary. The source dictionary.
'@param layer Byte. A ShowHideWorksheetLayer value. Which layer to write.
'@param sheetName String. The base sheet name.
'@param sourceSh Worksheet. The sheet of that layer in the source workbook.
Private Sub SaveOneShowHideLayer(ByVal outStore As ShowHideStore, _
                                 ByVal dict As LLdictionary, _
                                 ByVal layer As Byte, _
                                 ByVal sheetName As String, _
                                 ByVal sourceSh As Worksheet)

    Dim entries As ShowHide
    Dim layout As ShowHideLayout
    Dim tableName As String

    Set entries = ShowHide.Create(dict, layer, sheetName)
    If entries.EntryCount = 0 Then Exit Sub

    'The layout is handed the keys, because reading the real size of a hidden
    'column means showing it for an instant and the source sheets are protected.
    tableName = BaseTableNameOf(sourceSh)
    Set layout = ShowHideLayout.Create(sourceSh, layer, PasswordsObject(), tableName)

    ImportCurrentShowHide entries
    entries.Adopt layout

    outStore.Save entries, layout
End Sub

'@label:companion-sheet
'@sub-title Find a prefixed companion sheet of the source workbook
'@param sheetName String. The full sheet name, prefix included.
'@return Worksheet. The sheet, or Nothing when the workbook has none.
Private Function CompanionSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set CompanionSheet = this.sourceWkb.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:base-table-name-of
'@sub-title The table name a sheet's PRINTSTART anchor is named after
'@details
'A printed companion stores its table name with the print_ prefix in front and
'the anchor is named after the base one, so the prefix comes off here.
'@param sh Worksheet. The sheet to read.
'@return String. The table name with no print_ prefix.
Private Function BaseTableNameOf(ByVal sh As Worksheet) As String
    Dim tableName As String

    tableName = SheetTagOf(sh, "table_name")
    If InStr(1, tableName, PRINTPREFIX, vbTextCompare) = 1 Then
        tableName = Mid$(tableName, Len(PRINTPREFIX) + 1)
    End If

    BaseTableNameOf = tableName
End Function

'@label:import-current-show-hide
'@sub-title Fold the user's live choices into an entry list
'@details
'Reads the show/hide store of the source workbook. A workbook that has never
'had the form opened carries an empty table, and the entry list then keeps the
'state the dictionary authored.
'@param entries ShowHide. The entry list to update.
Private Sub ImportCurrentShowHide(ByVal entries As ShowHide)
    Dim showHideSh As Worksheet
    Dim srcStore As ShowHideStore

    Set showHideSh = InternalSheetOf(this.sourceWkb, SHOWHIDESHEET, SHOWHIDESHEET_OLD)
    If showHideSh Is Nothing Then Exit Sub

    Set srcStore = ShowHideStore.CreateForRead(showHideSh)
    If Not srcStore.HasTable Then Exit Sub

    srcStore.Load entries
End Sub

'@label:add-analysis-sheets
'@sub-title Add analysis worksheets to the output workbook
'@details
'Copies values and formats from the four analysis sheets (univariate,
'time series, spatial, spatio-temporal) in the source workbook to new
'sheets in the output workbook. Removes gridlines, then asks AnaTabIds to
'recreate the named ranges and the charts of each sheet.
'
'WHAT THE ERROR TRAP COVERS
'-------------------------------------------------------------------------------
'The trap covers the clipboard steps alone. It used to cover the chart writing
'too, and the chart writing raised on its first statement for every scope, so
'the export reported success and delivered analysis sheets with no named ranges
'and no charts. Each scope is written under its own trap now, the failures are
'collected, and the export refuses to report success while one is held.
'@param outwb Workbook. The output workbook.
'@param afterSheetName String. Name of the sheet to insert after.
'@throws ProjectError.SomethingWentWrong When a scope failed to write its
'charts.
Private Sub AddAnalysisSheets(ByVal outwb As Workbook, ByVal afterSheetName As String)

    Dim wkbNames As HiddenNames
    Dim sheetNamesList As BetterArray
    Dim scopesList As BetterArray
    Dim failuresList As BetterArray
    Dim namObj As AnaTabIds
    Dim outsh As Worksheet
    Dim inpsh As Worksheet
    Dim sheetName As String
    Dim scope As Byte
    Dim counter As Long
    Dim prevSheetName As String
    Dim graphsErrNumber As Long
    Dim graphsErrText As String
    Dim failureText As String

    Set wkbNames = HiddenNames.Create(this.sourceWkb)
    Set sheetNamesList = New BetterArray
    Set scopesList = New BetterArray
    Set failuresList = New BetterArray
    prevSheetName = afterSheetName

    sheetNamesList.Push wkbNames.ValueAsString("RNG_UASheet"), _
                        wkbNames.ValueAsString("RNG_TSSheet"), _
                        wkbNames.ValueAsString("RNG_SPSheet"), _
                        wkbNames.ValueAsString("RNG_SPTSheet")

    scopesList.Push AnalysisScopeNormal, AnalysisScopeTimeSeries, _
                    AnalysisScopeSpatial, AnalysisScopeSpatioTemporal

    Set namObj = AnaTabIds.Create(this.sourceWkb.Worksheets(ANANAMESSHEET))

    For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound
        sheetName = sheetNamesList.Item(counter)
        scope = scopesList.Item(counter)

        outwb.Worksheets.Add(After:=outwb.Worksheets(prevSheetName)).Name = sheetName
        prevSheetName = sheetName
        Set outsh = outwb.Worksheets(sheetName)

        ' Two of the four analysis sheets are built only when the setup asks
        ' for them, so a missing source sheet is a scope with nothing to copy.
        Set inpsh = Nothing
        On Error Resume Next
        Set inpsh = this.sourceWkb.Worksheets(sheetName)
        On Error GoTo 0

        If Not inpsh Is Nothing Then
            ' The clipboard steps, and only those
            On Error Resume Next
            Application.CutCopyMode = False
            inpsh.Cells.Copy
            outsh.Cells.PasteSpecial xlPasteValues

            Application.CutCopyMode = False
            inpsh.Cells.Copy
            outsh.Cells.PasteSpecial xlPasteFormats
            On Error GoTo 0

            RemoveGridLines outsh

            ' The names and the charts of this scope. A failure here is held
            ' and reported after every scope has had its turn, so one broken
            ' scope does not cost the others their values and formats.
            graphsErrNumber = 0
            On Error Resume Next
            namObj.WriteGraphs outsh, scope, inpsh
            graphsErrNumber = Err.Number
            graphsErrText = Err.Description
            On Error GoTo 0

            If graphsErrNumber <> 0 Then
                failuresList.Push sheetName & " (" & graphsErrNumber & _
                                  ": " & graphsErrText & ")"
            End If
        End If
    Next counter

    If failuresList.Length = 0 Then Exit Sub

    failureText = vbNullString
    For counter = failuresList.LowerBound To failuresList.UpperBound
        If LenB(failureText) > 0 Then failureText = failureText & ", "
        failureText = failureText & failuresList.Item(counter)
    Next counter

    ThrowError ProjectError.SomethingWentWrong, _
               "The analysis charts could not be written on: " & failureText
End Sub


'@section Private - Data Export
'===============================================================================

'@label:add-data-sheets
'@sub-title Add data to each output sheet from source HList/VList worksheets
'@details
'Iterates over the unique sheet names in the export dictionary, reads data
'from the corresponding source sheet, and writes it to the output sheet.
'For HList sheets, delegates to CustomTable.Export. For VList sheets,
'delegates to ExportVListData. Applies formatting to each output sheet.
'@param outwb Workbook. The output workbook.
'@param expDict LLdictionary. The export dictionary with variable assignments.
'@param useLabels Boolean. When True, writes label row above data headers.
'@param startLine Long. Data start row (2 when labels are on row 1).
'@param useFilters Boolean. When True, reads from filtered companion sheets.
Private Sub AddDataSheets(ByVal outwb As Workbook, ByVal expDict As LLdictionary, _
                          ByVal useLabels As Boolean, ByVal startLine As Long, _
                          ByVal useFilters As Boolean)

    Dim sheetNames As BetterArray
    Dim headersList As BetterArray
    Dim labelsList As BetterArray
    Dim sheetName As String
    Dim inpsh As Worksheet
    Dim inpshNames As HiddenNames
    Dim outsh As Worksheet
    Dim lo As ListObject
    Dim csTab As CustomTable
    Dim pass As Passwords
    Dim counter As Long

    Set sheetNames = expDict.UniqueValues("sheet name")
    Set pass = PasswordsObject()

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)
        Set inpsh = this.sourceWkb.Worksheets(sheetName)
        ' Build the sheet metadata once and read every tag from the same instance
        Set inpshNames = HiddenNames.Create(inpsh)
        Set headersList = expDict.Data.FilterData("sheet name", sheetName, _
                                                  "variable name", matchCase:=False)
        Set outsh = outwb.Worksheets(sheetName)

        Select Case inpshNames.ValueAsString("sheet_type")

        Case "HList"
            ' Choose between filtered or raw data
            If useFilters Then
                Set lo = FilteredTableOf(inpsh, inpshNames)
            Else
                Set lo = inpsh.ListObjects(1)
            End If

            Set csTab = CustomTable.Create(lo)
            If Not pass Is Nothing Then pass.UnProtect inpsh.Name
            csTab.Export outsh, headersList, startLine
            If Not pass Is Nothing Then pass.Protect inpsh.Name

            ' Add label row if configured. The export skips a header the table
            ' has no column for and leaves its output column where it was, so
            ' the labels of the skipped variables come out of the list before
            ' the row is written. Without that, one variable that is in the
            ' dictionary and absent from the ListObject put every label after
            ' it one column left of its data, for the rest of the sheet.
            If useLabels Then
                Set labelsList = LabelsForWrittenColumns(expDict, sheetName, _
                                                         headersList, _
                                                         csTab.ExportColumnsNotFound)
                labelsList.ToExcelRange outsh.Cells(1, 1), TransposeValues:=True
            End If

        Case "VList"
            ExportVListData inpsh, outsh, headersList, useLabels

        End Select

        FormatOutputSheet outsh
    Next counter
End Sub

'@label:filtered-table-of
'@sub-title The filtered companion table of one data entry sheet
'@details
'The companion holds the rows the filters leave visible, and the export reads it
'when the user ticks the filtered box. Three things have to be there: the name
'in the sheet store, the sheet it names, and the table on that sheet. All three
'used to be read in one chained line, so a workbook missing any of them raised a
'bare Excel number and the box the user got named neither the sheet nor the
'piece that was missing.
'@param inpsh Worksheet. The data entry sheet being exported.
'@param inpshNames HiddenNames. The hidden name store of that sheet.
'@return ListObject. The companion table.
'@throws ProjectError.ElementNotFound When any of the three is absent.
Private Function FilteredTableOf(ByVal inpsh As Worksheet, _
                                 ByVal inpshNames As HiddenNames) As ListObject

    Dim filtsh As Worksheet
    Dim filteredName As String

    filteredName = inpshNames.ValueAsString("filtered_sheet")
    If LenB(filteredName) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "The sheet " & inpsh.Name & " carries no filtered_sheet name, " & _
                   "so the export has no filtered data to read"
    End If

    On Error Resume Next
    Set filtsh = this.sourceWkb.Worksheets(filteredName)
    On Error GoTo 0

    If filtsh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "The workbook carries no worksheet named " & filteredName & _
                   ", which is the filtered companion of " & inpsh.Name
    End If

    On Error Resume Next
    Set FilteredTableOf = filtsh.ListObjects(1)
    On Error GoTo 0

    If FilteredTableOf Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "The filtered companion " & filteredName & " carries no table"
    End If
End Function

'@label:labels-for-written-columns
'@sub-title The label row that lines up with the columns an export wrote
'@details
'The dictionary answers one label per variable, in the same order as the
'variable names, so the two lists match until the export skips one. Dropping the
'label of every skipped variable is what keeps the row over its data.
'@param expDict LLdictionary. The export dictionary.
'@param sheetName String. The sheet the labels belong to.
'@param headersList BetterArray. The variable names the export was asked for.
'@param skipped BetterArray. The variable names the export could not write.
'@return BetterArray. One label per column written, in column order.
Private Function LabelsForWrittenColumns(ByVal expDict As LLdictionary, _
                                         ByVal sheetName As String, _
                                         ByVal headersList As BetterArray, _
                                         ByVal skipped As BetterArray) As BetterArray

    Dim allLabels As BetterArray
    Dim written As BetterArray
    Dim counter As Long
    Dim labelIndex As Long

    Set allLabels = expDict.Data.FilterData("sheet name", sheetName, _
                                            "main label", matchCase:=False)
    Set LabelsForWrittenColumns = allLabels

    If skipped Is Nothing Then Exit Function
    If skipped.Length = 0 Then Exit Function

    Set written = New BetterArray
    written.LowerBound = 1

    For counter = headersList.LowerBound To headersList.UpperBound
        labelIndex = allLabels.LowerBound + (counter - headersList.LowerBound)
        If labelIndex > allLabels.UpperBound Then Exit For

        If Not skipped.Includes(CStr(headersList.Item(counter))) Then
            written.Push allLabels.Item(labelIndex)
        End If
    Next counter

    Set LabelsForWrittenColumns = written
End Function

'@label:export-vlist-data
'@sub-title Export VList data (variable/value pairs) to the output sheet
'@details
'Writes one row per variable: the variable name in column A, its label beside it
'when the header format asks for labels, then the value.
'
'THE THREE HEADER FORMATS ON A VLIST SHEET
'-------------------------------------------------------------------------------
'Owner decision, 2026-08-01: the shape is `variable | label | value`, and the
'label column is written in the two label formats alone.
'
'What it replaces was wrong in two ways at once. The data start ROW was fed into
'a column index, so the `variable names + variable labels` format wrote the pair
'into columns B and C and left column A empty; and in the same format the
'variable name was overwritten by the label, so the file carried no variable
'names at all. `LLImporter.ImportVListSheet` resolves each row through
'`currsh.Range(varName)`, so every row of such a file landed in "not imported".
'Column A now always holds the variable name, whatever the format.
'@param inpsh Worksheet. The source VList worksheet.
'@param outsh Worksheet. The output worksheet.
'@param headersList BetterArray. Variable names to export.
'@param useLabels Boolean. When True, writes a label column beside the names.
Private Sub ExportVListData(ByVal inpsh As Worksheet, ByVal outsh As Worksheet, _
                            ByVal headersList As BetterArray, _
                            ByVal useLabels As Boolean)

    Dim varName As String
    Dim varValue As String
    Dim varLabel As String
    Dim counter As Long
    Dim outputRow As Long
    Dim valueColumn As Long

    outsh.Cells.Clear

    valueColumn = IIf(useLabels, 3, 2)

    outsh.Cells(1, 1).Value = "variable"
    If useLabels Then outsh.Cells(1, 2).Value = "label"
    outsh.Cells(1, valueColumn).Value = "value"

    outputRow = 2

    For counter = headersList.LowerBound To headersList.UpperBound
        varName = headersList.Item(counter)
        varValue = vbNullString
        varLabel = vbNullString

        On Error Resume Next
        varValue = CStr(inpsh.Range(varName).Value)
        varLabel = CStr(inpsh.Range(varName).Offset(, -1).Value)
        On Error GoTo 0

        outsh.Cells(outputRow, 1).Value = varName
        If useLabels Then outsh.Cells(outputRow, 2).Value = varLabel
        outsh.Cells(outputRow, valueColumn).Value = varValue

        outputRow = outputRow + 1
    Next counter
End Sub


'@section Private - Formatting Helpers
'===============================================================================

'@label:format-output-sheet
'@sub-title Apply standard formatting to an output data sheet
'@details
'Sets row heights, column widths, font, header styling, and auto-fits
'columns. Wrapped in error handling to prevent failures on protected
'or special sheets.
'@param outsh Worksheet. The worksheet to format.
Private Sub FormatOutputSheet(ByVal outsh As Worksheet)
    On Error Resume Next
    With outsh
        .Cells.EntireRow.RowHeight = 14
        .Cells.EntireColumn.ColumnWidth = 35
        .Cells.Font.Name = "Consolas"
        .Cells.Font.Size = 9
        .Rows(1).Font.Size = 10
        .Rows(1).RowHeight = 25
        .Rows(1).Interior.Color = RGB(188, 212, 230)
        .Rows(1).VerticalAlignment = xlVAlignCenter
        .Rows(1).Font.Bold = True
        .Cells.EntireColumn.AutoFit
    End With
    On Error GoTo 0
End Sub

'@label:remove-gridlines
'@sub-title Remove gridlines from an analysis output sheet
'@param sh Worksheet. The worksheet to update.
Private Sub RemoveGridLines(ByVal sh As Worksheet)
    Dim View As WorksheetView

    On Error Resume Next
    For Each View In sh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = sh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = False
            Exit Sub
        End If
    Next View
    On Error GoTo 0
End Sub


'@section Private - Error Handling
'===============================================================================

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@details
'Wrapper around Err.Raise that standardises the source to CLASS_NAME.
'@param errNumber ProjectError. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
