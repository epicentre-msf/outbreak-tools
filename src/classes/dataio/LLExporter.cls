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
'Exports linelist data into standalone workbooks for distribution. Supports
'four export modes: custom filtered exports (ExportCustom), full migration
'exports (ExportMigration), analysis-only exports (ExportAnalysis), and
'geobase exports (ExportGeo). Each mode builds a temporary output workbook,
'populates it with the relevant data, metadata, and formatting, then saves
'and closes it.
'@depends LLdictionary, LLVariables, LLExport, Passwords, CustomTable
'@depends DataSheet, AnaTabIds, LLGeo, LLChoices, HiddenNames, LLTranslation
'@depends ShowHide, ShowHideLayout, ShowHideStore, BetterArray, Checking


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
Private Const TEMPSHEET As String = "temp__"
Private Const EXPORTSHEET As String = "Exports"
Private Const ANANAMESSHEET As String = "ana_tabnames__"
Private Const SHOWHIDESHEET As String = "show_hide__"
Private Const GEOSHEET As String = "Geo"

' Prefixes of the two companion sheets a data entry sheet can have
Private Const PRINTPREFIX As String = "print_"
Private Const CRFPREFIX As String = "crf_"

' P-codes tag list for removal from custom exports
Private Const PCODES_GEO_PCODE_ADM1 As String = "geo_pcode_adm1"
Private Const PCODES_GEO_PCODE_ADM2 As String = "geo_pcode_adm2"
Private Const PCODES_GEO_PCODE_ADM3 As String = "geo_pcode_adm3"
Private Const PCODES_GEO_PCODE_ADM4 As String = "geo_pcode_adm4"
Private Const PCODES_GEO_CONCAT_ADM1 As String = "geo_concat_adm1"
Private Const PCODES_GEO_CONCAT_ADM2 As String = "geo_concat_adm2"
Private Const PCODES_GEO_CONCAT_ADM3 As String = "geo_concat_adm3"
Private Const PCODES_GEO_CONCAT_ADM4 As String = "geo_concat_adm4"


'@section State
'===============================================================================

Private Type TLLExporter
    sourceWkb As Workbook
    lastPassword As String
    activeWkb As Workbook
    sealed As Boolean
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

'@label:excluded-variables-for
'@sub-title The variables a numbered export leaves out
'@details
'The selection an export makes before it writes anything: the geo p-code and geo
'concat variables when the export row turns p-codes off, and the personal
'identifiers when it turns those off. Reading it costs three column reads and
'writes nothing, so a caller can ask what a file will be missing before asking
'for the file.
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

    ExportCustom = vbNullString
    this.lastPassword = vbNullString

    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()
    Set dict = SourceDictionary(nbExp)
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
    ' written True here, so every custom export promised a show_hide__ sheet it
    ' had no copy of.
    withMetadata = (expObj.ColumnValue(exportNumber, "export metadata sheets") = "yes")
    If withMetadata Then
        AddMetadataSheets outwb, dict, nbExp
        AddCustomDropdowns outwb
        AddDictionary outwb, dict, nbExp, expName, excluded
        AddMetadataTags outwb, ExportScopeCustom, exportNumber, _
                        includeShowHide:=True, keepLabels:=True
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

    ' Delete the placeholder first sheet
    On Error Resume Next
    outwb.Worksheets(firstSheetName).Delete
    On Error GoTo 0

    ' Export data to each sheet
    headerFormat = expObj.ColumnValue(exportNumber, "header format")
    useLabels = (headerFormat = "variables labels") Or _
                (headerFormat = "variable names + variable labels")
    startLine = IIf(headerFormat = "variable names + variable labels", 2, 1)

    AddDataSheets outwb, expDict, useLabels, startLine, useFilters

    ' Build file path and save
    fileName = expObj.ExportFileName(exportNumber, dict, pass)
    fileExt = expObj.ColumnValue(exportNumber, "file format")
    If LenB(fileExt) = 0 Then fileExt = "xlsx"
    fileFormat = IIf(fileExt = "xlsb", xlExcel12, xlOpenXMLWorkbook)
    filePath = BuildFilePath(folderPath, fileName, fileExt)

    addPassword = (expObj.ColumnValue(exportNumber, "password") = "yes")
    If addPassword Then this.lastPassword = pass.Value("privatekey")

    SaveWorkbook outwb, filePath, fileFormat, this.lastPassword
    ExportCustom = filePath
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

    ExportMigration = vbNullString
    this.lastPassword = vbNullString

    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()
    Set dict = SourceDictionary(nbExp)
    Set pass = PasswordsObject()

    ' A migration carries the whole dictionary, so it excludes nothing.
    Set excluded = New BetterArray

    ' Write full dictionary to temp sheet (no filtering)
    PrepareTempDictionary dict, vbNullString, True, excluded

    Set expDict = ExportDictionaryFromTemp(nbExp)

    ' Create output workbook
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add metadata sheets
    AddMetadataSheets outwb, dict, nbExp
    AddCustomDropdowns outwb
    AddDictionary outwb, dict, nbExp, "__all__", excluded
    AddMetadataTags outwb, ExportScopeMigration, 0, includeShowHide, keepLabels

    ' Add show/hide with column widths if requested
    If includeShowHide Then AddShowHide outwb

    prevSheetName = METADATASHEET

    ' Add data worksheets
    Set sheetNames = expDict.UniqueValues("sheet name")
    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        outwb.Worksheets.Add(After:=outwb.Worksheets(prevSheetName)).Name = _
            sheetNames.Item(counter)
        prevSheetName = sheetNames.Item(counter)
    Next counter

    ' Delete the placeholder first sheet
    On Error Resume Next
    outwb.Worksheets(firstSheetName).Delete
    On Error GoTo 0

    ' Export all data (no filtering, no labels)
    AddDataSheets outwb, expDict, False, 1, False

    ' Save as xlsx (no password for migration)
    fileName = expObj.ExportFileName(0, dict, pass, exportAll:=True)
    filePath = BuildFilePath(folderPath, fileName, "xlsx")
    SaveWorkbook outwb, filePath, xlOpenXMLWorkbook, vbNullString

    ExportMigration = filePath
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

    ExportAnalysis = vbNullString
    this.lastPassword = vbNullString

    Set expObj = ExportObject()
    nbExp = expObj.NumberOfExports()
    Set dict = SourceDictionary(nbExp)
    Set pass = PasswordsObject()

    ' Create output workbook
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add analysis sheets
    AddAnalysisSheets outwb, firstSheetName

    ' Delete the placeholder first sheet
    On Error Resume Next
    outwb.Worksheets(firstSheetName).Delete
    On Error GoTo 0

    ' Save
    fileName = expObj.ExportFileName(0, dict, pass, exportAll:=True)
    filePath = BuildFilePath(folderPath, fileName & "_analysis", "xlsx")
    SaveWorkbook outwb, filePath, xlOpenXMLWorkbook, vbNullString

    ExportAnalysis = filePath
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

    ExportGeo = vbNullString

    Set geo = GeoObject()
    Set outwb = CreateOutputWorkbook()

    geobaseName = geo.GeoBaseName()
    If LenB(geobaseName) = 0 Then
        geobaseName = "OUTBREAK-TOOLS-GEOBASE-UNKNOWN"
    End If

    geobaseName = geobaseName & "_export" & Format$(Now, "yyyymmdd")
    If onlyHistoric Then geobaseName = geobaseName & "_historic"

    filePath = BuildFilePath(folderPath, geobaseName, "xlsx")

    ' Use LLGeo.ExportToFile to write the geobase data
    geo.ExportToFile outwb, onlyHistoric

    SaveWorkbook outwb, filePath, xlOpenXMLWorkbook, vbNullString
    ExportGeo = filePath
End Function

'@label:close-all
'@sub-title Close all output workbooks that may still be open
'@details
'Closes any output workbook left open after an export error, preventing
'orphaned minimized windows. Safe to call even when no workbook is open.
'@export
Public Sub CloseAll()
    On Error Resume Next
    If Not this.activeWkb Is Nothing Then
        this.activeWkb.Close savechanges:=False
        Set this.activeWkb = Nothing
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
Private Function PasswordsObject() As Passwords
    Set PasswordsObject = Passwords.Create(this.sourceWkb.Worksheets(PASSWORDSHEET))
End Function

'@label:geo-object
'@sub-title Get the geo object from the source workbook
'@return LLGeo. Geo instance from the Geo sheet.
Private Function GeoObject() As LLGeo
    Set GeoObject = LLGeo.Create(this.sourceWkb.Worksheets(GEOSHEET))
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
'Reads the two "include" flags of the export row and answers the names of the
'variables the export drops: the geo p-code and geo concat variables when
'"include p-codes" is off, and every variable flagged as a personal identifier
'when "include personal identifiers" is off. A variable that answers both is
'listed twice, which costs a lookup and nothing else.
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
    Dim dropPcodes As Boolean
    Dim dropIdentifiers As Boolean
    Dim counter As Long

    Set excluded = New BetterArray
    Set ExcludedVariables = excluded

    dropPcodes = (expObj.ColumnValue(exportNumber, "include p-codes") <> "yes")
    dropIdentifiers = (expObj.ColumnValue(exportNumber, _
                                          "include personal identifiers") <> "yes")

    If Not dropPcodes And Not dropIdentifiers Then Exit Function

    varNames = ColumnValues(dict.DataRange("variable name"))

    If dropPcodes Then
        Set pcodesTagList = New BetterArray
        pcodesTagList.Push PCODES_GEO_PCODE_ADM1, PCODES_GEO_PCODE_ADM2, _
                           PCODES_GEO_PCODE_ADM3, PCODES_GEO_PCODE_ADM4, _
                           PCODES_GEO_CONCAT_ADM1, PCODES_GEO_CONCAT_ADM2, _
                           PCODES_GEO_CONCAT_ADM3, PCODES_GEO_CONCAT_ADM4

        mainLabels = ColumnValues(dict.DataRange("main label"))
        For counter = LBound(varNames) To UBound(varNames)
            If pcodesTagList.Includes(CStr(mainLabels(counter))) Then
                excluded.Push CStr(varNames(counter))
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
'memory. Writes the result to the temp__ sheet and adds a blue font tag
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

    Set tempsh = this.sourceWkb.Worksheets(TEMPSHEET)

    If fullExport Then
        Set dictData = New BetterArray
        dictData.FromExcelRange dict.DataRange(includeHeaders:=True)
    Else
        Set dictData = dict.Data.FilterData( _
            varName:=expName, criteriaName:="<>", _
            returnedColumnName:="__all__", includeHeaders:=True)
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
'@return LLdictionary. Dictionary built from the temp__ sheet.
Private Function ExportDictionaryFromTemp(ByVal nbExp As Long) As LLdictionary
    Set ExportDictionaryFromTemp = LLdictionary.Create( _
        this.sourceWkb.Worksheets(TEMPSHEET), 1, 1, nbExp)
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
    Dim translations As LLTranslation

    Set sheetNames = New BetterArray
    sheetNames.Push METADATASHEET, CHOICESSHEET

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)
        Set sh = this.sourceWkb.Worksheets(sheetName)
        Set csData = DataSheet.Create(sh, 1, 1)
        csData.Export toWkb:=outwb, Hide:=xlSheetVisible
    Next counter

    ' Export dictionary translations via LLTranslation
    Set translations = LLTranslation.Create( _
        this.sourceWkb.Worksheets(LINELISTTRANSLATIONSHEET))
    translations.ExportDictionary outwb
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
'Reads custom dropdown ListObjects from the source workbook custom
'dropdown sheet and adds each one to the output Choices sheet. Silently
'exits when no custom dropdown sheet exists.
'@param outwb Workbook. The output workbook.
Private Sub AddCustomDropdowns(ByVal outwb As Workbook)
    Dim wkbNames As HiddenNames
    Dim choi As LLChoices
    Dim lo As ListObject
    Dim loValues As BetterArray
    Dim sh As Worksheet
    Dim choiceName As String

    Set wkbNames = HiddenNames.Create(this.sourceWkb)
    Set choi = LLChoices.Create(outwb.Worksheets(CHOICESSHEET), 1, 1)
    Set loValues = New BetterArray

    On Error Resume Next
    Set sh = this.sourceWkb.Worksheets(wkbNames.ValueAsString("RNG_CustomDrop"))
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    For Each lo In sh.ListObjects
        loValues.Clear
        loValues.FromExcelRange lo.Range.Cells(2, 1), _
            DetectLastRow:=True, DetectLastColumn:=False
        choiceName = "__choice_custom_" & lo.Range.Cells(1, 1).Value
        choi.AddChoice choiceName, loValues
    Next lo
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
Private Sub AddMetadataTags(ByVal outwb As Workbook, ByVal scope As Byte, _
                            ByVal exportNumber As Long, _
                            ByVal includeShowHide As Boolean, _
                            ByVal keepLabels As Boolean)

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
End Sub

'@label:add-show-hide
'@sub-title Export the show/hide choices with their sizes to the output workbook
'@details
'Builds a show_hide__ sheet in the output workbook and asks ShowHideStore to
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

    Set showHideSh = Nothing
    On Error Resume Next
    Set showHideSh = this.sourceWkb.Worksheets(SHOWHIDESHEET)
    On Error GoTo 0

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
        Set headersList = expDict.Data.FilterData("sheet name", sheetName, "variable name")
        Set outsh = outwb.Worksheets(sheetName)

        Select Case inpshNames.ValueAsString("sheet_type")

        Case "HList"
            ' Choose between filtered or raw data
            If useFilters Then
                Set lo = this.sourceWkb.Worksheets( _
                    inpshNames.ValueAsString("filtered_sheet")).ListObjects(1)
            Else
                Set lo = inpsh.ListObjects(1)
            End If

            Set csTab = CustomTable.Create(lo)
            pass.UnProtect inpsh.Name
            csTab.Export outsh, headersList, startLine
            pass.Protect inpsh.Name

            ' Add label row if configured
            If useLabels Then
                Set labelsList = expDict.Data.FilterData("sheet name", sheetName, "main label")
                labelsList.ToExcelRange outsh.Cells(1, 1), TransposeValues:=True
            End If

        Case "VList"
            ExportVListData inpsh, outsh, headersList, useLabels

        End Select

        FormatOutputSheet outsh
    Next counter
End Sub

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
