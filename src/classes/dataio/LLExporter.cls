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
Implements ILLExporter

'@class LLExporter
'@description
'Exports linelist data into standalone workbooks for distribution. Supports
'four export modes: custom filtered exports (ExportCustom), full migration
'exports (ExportMigration), analysis-only exports (ExportAnalysis), and
'geobase exports (ExportGeo). Each mode builds a temporary output workbook,
'populates it with the relevant data, metadata, and formatting, then saves
'and closes it. Consumers interact through the ILLExporter interface.
'@depends ILLdictionary, ILLExport, IPasswords, ICustomTable
'@depends IDataSheet, IShowHideManager, IAnaTabIds, ILLGeo, ILLChoices, IHiddenNames


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
End Type

Private this As TLLExporter


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create an exporter from the source linelist workbook
'@details
'Factory method that validates the workbook is not Nothing, then constructs
'and returns a new LLExporter instance through the ILLExporter interface.
'Called on the predeclared instance.
'@param sourceWkb Workbook. The linelist workbook containing the data to export.
'@return ILLExporter. A fully initialised instance ready for export operations.
'@throws ProjectError.ObjectNotInitialized When the workbook is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook) As ILLExporter

    If sourceWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "LLExporter requires a source workbook"
    End If

    With New LLExporter
        Set .SourceWorkbook = sourceWkb
        Set Create = .Self
    End With
End Function


'@section Public Accessors
'===============================================================================

'@label:self
'@prop-title Current object instance
'@return ILLExporter. The current instance cast to the interface.
Public Property Get Self() As ILLExporter
    Set Self = Me
End Property

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
    Set this.sourceWkb = wb
End Property

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
    Dim expObj As ILLExport
    Dim dict As ILLdictionary
    Dim pass As IPasswords
    Dim expDict As ILLdictionary
    Dim sheetNames As BetterArray
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

    ' Remove p-codes and personal identifiers if configured
    RemovePCodes dict, expObj, exportNumber, expName
    RemovePersonalIdentifiers dict, expObj, exportNumber, expName

    ' Filter dictionary for this export and write to temp sheet
    PrepareTempDictionary dict, expName, False
    Set expDict = ExportDictionaryFromTemp(nbExp)

    ' Sort export dictionary by table name and export column
    SortExportDictionary expDict, expName

    ' Create output workbook
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add metadata if configured
    withMetadata = (expObj.ColumnValue(exportNumber, "export metadata sheets") = "yes")
    If withMetadata Then
        AddMetadataSheets outwb, dict, nbExp
        AddCustomDropdowns outwb
        AddDictionary outwb, dict, nbExp, expName
        AddMetadataTags outwb, ExportScopeCustom, exportNumber, _
                        includeShowHide:=True, keepLabels:=True
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
    Dim expObj As ILLExport
    Dim dict As ILLdictionary
    Dim pass As IPasswords
    Dim expDict As ILLdictionary
    Dim sheetNames As BetterArray
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

    ' Write full dictionary to temp sheet (no filtering)
    PrepareTempDictionary dict, vbNullString, True

    Set expDict = ExportDictionaryFromTemp(nbExp)

    ' Create output workbook
    Set outwb = CreateOutputWorkbook()
    firstSheetName = outwb.Worksheets(1).Name

    ' Add metadata sheets
    AddMetadataSheets outwb, dict, nbExp
    AddCustomDropdowns outwb
    AddDictionary outwb, dict, nbExp, "__all__"
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
    Dim expObj As ILLExport
    Dim dict As ILLdictionary
    Dim pass As IPasswords
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
    Dim geo As ILLGeo
    Dim geobaseName As String
    Dim filePath As String

    ExportGeo = vbNullString

    Set geo = GeoObject()
    Set outwb = CreateOutputWorkbook()

    geobaseName = geo.GeoNames("geobase")
    If LenB(geobaseName) = 0 Or geobaseName = "geobase" Then
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
Private Function SheetTagOf(ByVal sh As Worksheet) As String
    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)
    SheetTagOf = shHn.ValueAsString("sheet_type")
    If SheetTagOf = vbNullString Then SheetTagOf = CStr(sh.Cells(1, 3).Value)
End Function

'@section Private - Source Object Accessors
'===============================================================================

'@label:source-dictionary
'@sub-title Get a dictionary from the source workbook
'@param nbExp Long. Number of export columns to configure.
'@return ILLdictionary. Dictionary instance from the Dictionary sheet.
Private Function SourceDictionary(ByVal nbExp As Long) As ILLdictionary
    Set SourceDictionary = LLdictionary.Create( _
        this.sourceWkb.Worksheets(DICTIONARYSHEET), 1, 1, nbExp)
End Function

'@label:export-object
'@sub-title Get the export configuration from the source workbook
'@return ILLExport. Export configuration from the Exports sheet.
Private Function ExportObject() As ILLExport
    Set ExportObject = LLExport.Create(this.sourceWkb.Worksheets(EXPORTSHEET))
End Function

'@label:passwords-object
'@sub-title Get passwords from the source workbook
'@return IPasswords. Passwords instance from the __pass sheet.
Private Function PasswordsObject() As IPasswords
    Set PasswordsObject = Passwords.Create(this.sourceWkb.Worksheets(PASSWORDSHEET))
End Function

'@label:geo-object
'@sub-title Get the geo object from the source workbook
'@return ILLGeo. Geo instance from the Geo sheet.
Private Function GeoObject() As ILLGeo
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

'@label:remove-pcodes
'@sub-title Remove p-codes variables from a custom export if configured
'@details
'Checks the export configuration for "include p-codes". When set to "no",
'clears the export column cells for all dictionary rows whose main label
'matches a geo p-code or geo concat tag.
'@param dict ILLdictionary. The source dictionary.
'@param expObj ILLExport. The export configuration.
'@param exportNumber Long. Which export definition to check.
'@param expName String. The export column name (e.g. "export 1").
Private Sub RemovePCodes(ByVal dict As ILLdictionary, ByVal expObj As ILLExport, _
                         ByVal exportNumber As Long, ByVal expName As String)

    Dim exportPcodes As Boolean
    Dim pcodesTagList As BetterArray
    Dim expNumberRng As Range
    Dim mainLabRng As Range
    Dim counter As Long

    exportPcodes = (expObj.ColumnValue(exportNumber, "include p-codes") = "yes")
    If exportPcodes Then Exit Sub

    Set pcodesTagList = New BetterArray
    pcodesTagList.Push PCODES_GEO_PCODE_ADM1, PCODES_GEO_PCODE_ADM2, _
                       PCODES_GEO_PCODE_ADM3, PCODES_GEO_PCODE_ADM4, _
                       PCODES_GEO_CONCAT_ADM1, PCODES_GEO_CONCAT_ADM2, _
                       PCODES_GEO_CONCAT_ADM3, PCODES_GEO_CONCAT_ADM4

    Set expNumberRng = dict.DataRange(expName)
    Set mainLabRng = dict.DataRange("main label")

    For counter = 1 To mainLabRng.Rows.count
        If pcodesTagList.Includes(CStr(mainLabRng.Cells(counter, 1).Value)) Then
            expNumberRng.Cells(counter, 1).ClearContents
        End If
    Next counter
End Sub

'@label:remove-personal-identifiers
'@sub-title Remove personal identifiers from a custom export if configured
'@details
'Checks the export configuration for "include personal identifiers". When
'set to "no", clears the export column cells for all dictionary rows
'flagged as personal identifiers.
'@param dict ILLdictionary. The source dictionary.
'@param expObj ILLExport. The export configuration.
'@param exportNumber Long. Which export definition to check.
'@param expName String. The export column name.
Private Sub RemovePersonalIdentifiers(ByVal dict As ILLdictionary, ByVal expObj As ILLExport, _
                                      ByVal exportNumber As Long, ByVal expName As String)

    Dim exportPI As Boolean
    Dim expNumberRng As Range
    Dim identifierRng As Range
    Dim counter As Long

    exportPI = (expObj.ColumnValue(exportNumber, "include personal identifiers") = "yes")
    If exportPI Then Exit Sub

    Set expNumberRng = dict.DataRange(expName)
    Set identifierRng = dict.DataRange("personal identifier")

    For counter = 1 To identifierRng.Rows.count
        If CStr(identifierRng.Cells(counter, 1).Value) = "yes" Then
            expNumberRng.Cells(counter, 1).ClearContents
        End If
    Next counter
End Sub

'@label:prepare-temp-dictionary
'@sub-title Write dictionary data to the temp sheet for export processing
'@details
'When fullExport is True, copies the entire dictionary. When False,
'filters on the export column name to include only rows assigned to that
'export. Writes the result to the temp__ sheet and adds a blue font tag
'to mark the dictionary as prepared.
'@param dict ILLdictionary. The source dictionary.
'@param expName String. The export column name to filter on, or vbNullString for full.
'@param fullExport Boolean. When True, copies the entire dictionary unfiltered.
Private Sub PrepareTempDictionary(ByVal dict As ILLdictionary, _
                                  ByVal expName As String, _
                                  ByVal fullExport As Boolean)

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
'@return ILLdictionary. Dictionary built from the temp__ sheet.
Private Function ExportDictionaryFromTemp(ByVal nbExp As Long) As ILLdictionary
    Set ExportDictionaryFromTemp = LLdictionary.Create( _
        this.sourceWkb.Worksheets(TEMPSHEET), 1, 1, nbExp)
End Function

'@label:sort-export-dictionary
'@sub-title Sort the export dictionary by table name and export column
'@param expDict ILLdictionary. The dictionary to sort.
'@param expName String. The export column name to use as secondary sort key.
Private Sub SortExportDictionary(ByVal expDict As ILLdictionary, ByVal expName As String)
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
'@param dict ILLdictionary. The source dictionary.
'@param nbExp Long. Number of export columns.
Private Sub AddMetadataSheets(ByVal outwb As Workbook, ByVal dict As ILLdictionary, _
                              ByVal nbExp As Long)

    Dim sheetNames As BetterArray
    Dim csData As IDataSheet
    Dim sh As Worksheet
    Dim counter As Long
    Dim sheetName As String
    Dim translations As ILLTranslation

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
'Exports the dictionary filtered by exportColumn, then marks the output
'dictionary as prepared with a blue font tag.
'@param outwb Workbook. The output workbook.
'@param dict ILLdictionary. The source dictionary.
'@param nbExp Long. Number of export columns.
'@param exportColumn String. The export column to filter on, or "__all__".
Private Sub AddDictionary(ByVal outwb As Workbook, ByVal dict As ILLdictionary, _
                          ByVal nbExp As Long, ByVal exportColumn As String)

    Dim outDict As ILLdictionary
    Dim outDictRng As Range

    dict.Export toWkb:=outwb, exportType:=exportColumn, _
               addListObject:=False, Hide:=xlSheetVisible

    ' Mark the output dictionary as prepared
    Set outDict = LLdictionary.Create(outwb.Worksheets(DICTIONARYSHEET), 1, 1, nbExp)
    Set outDictRng = outDict.DataRange()
    outDictRng.Cells(outDictRng.Rows.count + 1, 1).Font.Color = vbBlue
End Sub

'@label:add-custom-dropdowns
'@sub-title Add custom dropdowns to the Choices sheet in the output workbook
'@details
'Reads custom dropdown ListObjects from the source workbook custom
'dropdown sheet and adds each one to the output Choices sheet. Silently
'exits when no custom dropdown sheet exists.
'@param outwb Workbook. The output workbook.
Private Sub AddCustomDropdowns(ByVal outwb As Workbook)
    Dim wkbNames As IHiddenNames
    Dim choi As ILLChoices
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

    Dim csData As IDataSheet
    Dim varRng As Range
    Dim expName As String
    Dim rowOffset As Long
    Dim epiWeekStart As String

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

    ' Epiweek start value
    epiWeekStart = vbNullString
    On Error Resume Next
    epiWeekStart = CStr(this.sourceWkb.Names("RNG_EpiWeekStart").RefersToRange.Value)
    On Error GoTo 0

    varRng.Cells(rowOffset + 5, 1).Value = "epiweekstart"
    varRng.Cells(rowOffset + 5, 2).Value = epiWeekStart
End Sub

'@label:add-show-hide
'@sub-title Export show/hide state with column widths to the output workbook
'@details
'Creates a show_hide__ sheet in the output workbook with a ListObject
'containing layer, field_key, header_text, hidden_flag, and column_width
'columns. For each data sheet, creates a ShowHideManager per applicable
'layer (HList, Printed, VList) and exports the current visibility state.
'@param outwb Workbook. The output workbook.
Private Sub AddShowHide(ByVal outwb As Workbook)

    Dim outsh As Worksheet
    Dim dict As ILLdictionary
    Dim sheetNames As BetterArray
    Dim sheetName As String
    Dim sheetType As String
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim shm As IShowHideManager
    Dim counter As Long
    Dim nbExp As Long
    Dim loRng As Range

    nbExp = ExportObject().NumberOfExports()
    Set dict = SourceDictionary(nbExp)

    ' Create show_hide__ sheet in output workbook
    outwb.Worksheets.Add(outwb.Worksheets(outwb.Worksheets.count)).Name = SHOWHIDESHEET
    Set outsh = outwb.Worksheets(SHOWHIDESHEET)

    ' Create ListObject with 5 migration columns
    outsh.Cells(1, 1).Value = "layer"
    outsh.Cells(1, 2).Value = "field_key"
    outsh.Cells(1, 3).Value = "header_text"
    outsh.Cells(1, 4).Value = "hidden_flag"
    outsh.Cells(1, 5).Value = "column_width"

    Set loRng = outsh.Range(outsh.Cells(1, 1), outsh.Cells(1, 5))
    Set lo = outsh.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=loRng, _
        XlListObjectHasHeaders:=xlYes)
    lo.Name = "show_hide_migration"
    lo.TableStyle = "TableStyleLight4"

    ' For each data sheet, export show/hide per applicable layer
    Set sheetNames = dict.UniqueValues("sheet name")

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)

        ' Determine sheet type from the source workbook
        On Error Resume Next
        Set sh = this.sourceWkb.Worksheets(sheetName)
        On Error GoTo 0
        If sh Is Nothing Then GoTo NextSheet
        sheetType = SheetTagOf(sh)

        If sheetType = "HList" Then
            ' HList layer
            Set shm = ShowHideManager.Create(dict, ShowHideLayerHList, sheetName)
            ImportCurrentShowHide shm
            shm.ExportForMigration lo, sh

            ' Printed layer (companion print sheet)
            Dim printSh As Worksheet
            Set printSh = Nothing
            On Error Resume Next
            Set printSh = this.sourceWkb.Worksheets("print_" & sheetName)
            On Error GoTo 0

            If Not printSh Is Nothing Then
                Set shm = ShowHideManager.Create(dict, ShowHideLayerPrinted, sheetName)
                ImportCurrentShowHide shm
                shm.ExportForMigration lo, printSh
            End If

        ElseIf sheetType = "VList" Then
            ' VList layer
            Set shm = ShowHideManager.Create(dict, ShowHideLayerVList, sheetName)
            ImportCurrentShowHide shm
            shm.ExportForMigration lo, sh
        End If

        Set sh = Nothing
NextSheet:
    Next counter

    On Error Resume Next
    outsh.Cells.EntireColumn.AutoFit
    On Error GoTo 0
End Sub

'@label:import-current-show-hide
'@sub-title Import current user overrides from the show_hide__ ListObject
'@details
'Reads the show_hide__ sheet from the source workbook and imports the
'user visibility plan into the ShowHideManager, ensuring the exported
'show/hide reflects the user actual choices.
'@param shm IShowHideManager. The manager to populate.
Private Sub ImportCurrentShowHide(ByVal shm As IShowHideManager)
    Dim showHideSh As Worksheet
    Dim lo As ListObject

    Set showHideSh = Nothing
    On Error Resume Next
    Set showHideSh = this.sourceWkb.Worksheets(SHOWHIDESHEET)
    On Error GoTo 0

    If showHideSh Is Nothing Then Exit Sub
    If showHideSh.ListObjects.count = 0 Then Exit Sub

    Set lo = showHideSh.ListObjects(1)
    shm.ImportPlan lo
End Sub

'@label:add-analysis-sheets
'@sub-title Add analysis worksheets to the output workbook
'@details
'Copies values and formats from the four analysis sheets (univariate,
'time series, spatial, spatio-temporal) in the source workbook to new
'sheets in the output workbook. Removes gridlines and writes graph
'objects via AnaTabIds.
'@param outwb Workbook. The output workbook.
'@param afterSheetName String. Name of the sheet to insert after.
Private Sub AddAnalysisSheets(ByVal outwb As Workbook, ByVal afterSheetName As String)

    Dim wkbNames As IHiddenNames
    Dim sheetNamesList As BetterArray
    Dim scopesList As BetterArray
    Dim namObj As IAnaTabIds
    Dim outsh As Worksheet
    Dim sheetName As String
    Dim scope As Byte
    Dim counter As Long
    Dim prevSheetName As String

    Set wkbNames = HiddenNames.Create(this.sourceWkb)
    Set sheetNamesList = New BetterArray
    Set scopesList = New BetterArray
    prevSheetName = afterSheetName

    sheetNamesList.Push wkbNames.ValueAsString("RNG_UASheet"), _
                        wkbNames.ValueAsString("RNG_TSSheet"), _
                        wkbNames.ValueAsString("RNG_SPSheet"), _
                        wkbNames.ValueAsString("RNG_SPTSheet")

    scopesList.Push AnalysisIdsScopeNormal, AnalysisIdsScopeTimeSeries, _
                    AnalysisIdsScopeSpatial, AnalysisIdsScopeSpatioTemporal

    Set namObj = AnaTabIds.Create(this.sourceWkb.Worksheets(ANANAMESSHEET))

    For counter = sheetNamesList.LowerBound To sheetNamesList.UpperBound
        sheetName = sheetNamesList.Item(counter)
        scope = scopesList.Item(counter)

        outwb.Worksheets.Add(After:=outwb.Worksheets(prevSheetName)).Name = sheetName
        prevSheetName = sheetName
        Set outsh = outwb.Worksheets(sheetName)

        ' Copy values and formats from source analysis sheet
        On Error Resume Next
        Application.CutCopyMode = False
        this.sourceWkb.Worksheets(sheetName).Cells.Copy
        outsh.Cells.PasteSpecial xlPasteValues

        Application.CutCopyMode = False
        this.sourceWkb.Worksheets(sheetName).Cells.Copy
        outsh.Cells.PasteSpecial xlPasteFormats

        RemoveGridLines outsh
        namObj.WriteGraphs outsh, scope
        On Error GoTo 0
    Next counter
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
'@param expDict ILLdictionary. The export dictionary with variable assignments.
'@param useLabels Boolean. When True, writes label row above data headers.
'@param startLine Long. Data start row (2 when labels are on row 1).
'@param useFilters Boolean. When True, reads from filtered companion sheets.
Private Sub AddDataSheets(ByVal outwb As Workbook, ByVal expDict As ILLdictionary, _
                          ByVal useLabels As Boolean, ByVal startLine As Long, _
                          ByVal useFilters As Boolean)

    Dim sheetNames As BetterArray
    Dim headersList As BetterArray
    Dim labelsList As BetterArray
    Dim sheetName As String
    Dim inpsh As Worksheet
    Dim outsh As Worksheet
    Dim lo As ListObject
    Dim csTab As ICustomTable
    Dim pass As IPasswords
    Dim counter As Long

    Set sheetNames = expDict.UniqueValues("sheet name")
    Set pass = PasswordsObject()

    For counter = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(counter)
        Set inpsh = this.sourceWkb.Worksheets(sheetName)
        Set headersList = expDict.Data.FilterData("sheet name", sheetName, "variable name")
        Set outsh = outwb.Worksheets(sheetName)

        Select Case SheetTagOf(inpsh)

        Case "HList"
            ' Choose between filtered or raw data
            If useFilters Then
                Set lo = this.sourceWkb.Worksheets( _
                    CStr(inpsh.Cells(1, 5).Value)).ListObjects(1)
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
            ExportVListData inpsh, outsh, headersList, useLabels, startLine

        End Select

        FormatOutputSheet outsh
    Next counter
End Sub

'@label:export-vlist-data
'@sub-title Export VList data (variable/value pairs) to the output sheet
'@details
'Writes variable names and their values as rows. When useLabels is True,
'replaces variable names with their labels.
'@param inpsh Worksheet. The source VList worksheet.
'@param outsh Worksheet. The output worksheet.
'@param headersList BetterArray. Variable names to export.
'@param useLabels Boolean. When True, writes labels instead of variable names.
'@param startLine Long. Column offset for data start.
Private Sub ExportVListData(ByVal inpsh As Worksheet, ByVal outsh As Worksheet, _
                            ByVal headersList As BetterArray, _
                            ByVal useLabels As Boolean, ByVal startLine As Long)

    Dim varName As String
    Dim varValue As String
    Dim varLabel As String
    Dim counter As Long
    Dim headerCounter As Long
    Dim startCol As Long

    outsh.Cells.Clear
    startCol = IIf(startLine > 1, startLine, 1)

    outsh.Cells(1, startCol).Value = "variable"
    outsh.Cells(1, startCol + 1).Value = "value"

    headerCounter = 2

    For counter = headersList.LowerBound To headersList.UpperBound
        varName = headersList.Item(counter)
        varValue = vbNullString
        varLabel = vbNullString

        On Error Resume Next
        varValue = CStr(inpsh.Range(varName).Value)
        varLabel = CStr(inpsh.Range(varName).Offset(, -1).Value)
        On Error GoTo 0

        outsh.Cells(headerCounter, startCol).Value = varName
        outsh.Cells(headerCounter, startCol + 1).Value = varValue
        If useLabels Then outsh.Cells(headerCounter, startCol).Value = varLabel

        headerCounter = headerCounter + 1
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


'@section Interface Implementation
'===============================================================================
'@description Delegated members satisfying the ILLExporter contract.
'See the corresponding members above for full documentation.

Private Function ILLExporter_ExportCustom(ByVal exportNumber As Long, _
                                          ByVal folderPath As String, _
                                          Optional ByVal useFilters As Boolean = False) As String
    ILLExporter_ExportCustom = ExportCustom(exportNumber, folderPath, useFilters)
End Function

Private Function ILLExporter_ExportMigration(ByVal folderPath As String, _
                                             Optional ByVal includeShowHide As Boolean = False, _
                                             Optional ByVal keepLabels As Boolean = False) As String
    ILLExporter_ExportMigration = ExportMigration(folderPath, includeShowHide, keepLabels)
End Function

Private Function ILLExporter_ExportAnalysis(ByVal folderPath As String) As String
    ILLExporter_ExportAnalysis = ExportAnalysis(folderPath)
End Function

Private Function ILLExporter_ExportGeo(ByVal folderPath As String, _
                                       Optional ByVal onlyHistoric As Boolean = False) As String
    ILLExporter_ExportGeo = ExportGeo(folderPath, onlyHistoric)
End Function

Private Property Get ILLExporter_LastExportPassword() As String
    ILLExporter_LastExportPassword = LastExportPassword
End Property

Private Sub ILLExporter_CloseAll()
    CloseAll
End Sub
