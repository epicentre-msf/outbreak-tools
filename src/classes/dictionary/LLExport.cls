VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manage linelist and setup exports"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Manage linelist and setup exports")

'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument


'@depends Checking, CustomTable, LLVariables
Option Explicit

Implements ILLExport

Private Type TLLExport
    expData As IDataSheet
    exportSheet As Worksheet
    variables As ILLVariables
    checks As IChecking
    hasChecks As Boolean
    checkCounter As Long
    sanitizeList As Variant
End Type

Private Const COL_STATUS As String = "status"
Private Const COL_FILE_NAME As String = "file name"
Private Const COL_EXPORT_NUMBER As String = "export number"
Private Const COL_INCLUDE_IDS As String = "include personal identifiers"
Private Const STATUS_ACTIVE As String = "active"
Private Const DEFAULT_INCLUDE_IDS As String = "no"
Private Const TIMESTAMP_FORMAT As String = "yyyymmdd-HhNn"
Private Const CHECKING_CONTEXT As String = "LLExport checks"
Private Const CLASSNAME As String = "LLExport"

Private this As TLLExport

'@section Factory
'===============================================================================
'@label:Create
'@prop-title Create a unified export handler
'@details
'Creates a linelist export handler over the specification sheet provided by the
'caller. The implementation caches the underlying datasheet, worksheet, and
'sanitisation rules so subsequent calls can resolve filenames and table rows
'without repeatedly scanning the workbook.
'@param exportSheet Worksheet hosting the export specification table.
'@param startRow Optional; first row of the export table (defaults to 1).
'@param startColumn Optional; first column of the export table (defaults to 1).
'@return A fully initialised `ILLExport` instance.
'@throws ProjectError.ObjectNotInitialized when the worksheet is missing.
Public Function Create(ByVal exportSheet As Worksheet, _
                       Optional ByVal startRow As Long = 1, _
                       Optional ByVal startColumn As Long = 1) As ILLExport

    Dim exportData As IDataSheet

    If exportSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Worksheet is required to create " & CLASSNAME
    End If

    Set exportData = DataSheet.Create(exportSheet, startRow, startColumn)

    With New LLExport
        Set .InternalData = exportData
        Set .InternalWorksheet = exportSheet
        .ResetCaches
        Set Create = .Self()
    End With
End Function

'@label:SetInternalData
'@param exportData Datasheet wrapper over the specification table.
Public Property Set InternalData(ByVal exportData As IDataSheet)
    Set this.expData = exportData
End Property

'@label:SetInternalWorksheet
'@param exportSheet Worksheet that stores the export specification table.
Public Property Set InternalWorksheet(ByVal exportSheet As Worksheet)
    Set this.exportSheet = exportSheet
End Property

'@label:Self
'@prop-title Return the current instance as interface
'@return The object cast as `ILLExport`.
Public Property Get Self() As ILLExport
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'@label:Data
'@prop-title Underlying export datasheet
'@return Datasheet abstraction wrapping the export specification table.
Public Property Get Data() As IDataSheet
    EnsureDataReady
    Set Data = this.expData
End Property

'@label:Worksheet
'@prop-title Worksheet hosting the export specification table
'@return Worksheet reference used by the handler.
Public Property Get Worksheet() As Worksheet
    EnsureWorksheetReady
    Set Worksheet = this.exportSheet
End Property

'@label:NumberOfExports
'@prop-title Count the configured export rows
'@details Counts rows within the specification table excluding headers.
'@return Total number of export definitions currently present.
Public Property Get NumberOfExports() As Long
    Dim exportRange As Range

    EnsureDataReady

    Set exportRange = this.expData.DataRange(includeHeaders:=False)

    If exportRange Is Nothing Then
        NumberOfExports = 0
    Else
        NumberOfExports = exportRange.Rows.Count
    End If
End Property

'@label:HasCheckings
'@prop-title Whether informational checkings were captured
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:CheckingValues
'@prop-title Collected informational diagnostics
'@return Checking object containing log entries, if any.
Public Property Get CheckingValues() As IChecking
    If this.hasChecks Then
        Set CheckingValues = this.checks
    End If
End Property

'@section SpecManagement
'===============================================================================
'@label:AddRows
'@prop-title Append export definition rows
'@param rowCount Optional; number of rows to append (defaults to 1).
'@throws ProjectError.InvalidArgument when the requested row count is below 1.
Public Sub AddRows(Optional ByVal rowCount As Long = 1)
    Dim tableObj As ICustomTable
    Dim includeColumn As Range
    Dim idx As Long

    EnsureDataReady
    EnsureWorksheetReady

    If rowCount < 1 Then
        ThrowError ProjectError.InvalidArgument, "Row count must be greater than zero"
    End If

    Set tableObj = ExportTable()
    tableObj.AddRows nbRows:=CLng(rowCount), includeIds:=False

    Set includeColumn = tableObj.DataRange(COL_INCLUDE_IDS)
    If Not includeColumn Is Nothing Then
        For idx = 1 To includeColumn.Rows.Count
            If LenB(CStr(includeColumn.Cells(idx, 1).Value)) = 0 Then
                includeColumn.Cells(idx, 1).Value = DEFAULT_INCLUDE_IDS
            End If
        Next idx
    End If

    InvalidateCaches
End Sub

'@label:RemoveRows
'@prop-title Remove empty export rows
'@param totalCount Optional; number of rows to keep (0 trims trailing empty rows).
Public Sub RemoveRows(Optional ByVal totalCount As Long = 0)
    Dim tableObj As ICustomTable

    EnsureDataReady
    Set tableObj = ExportTable()
    tableObj.RemoveRows totalCount:=totalCount, includeIds:=False
    InvalidateCaches
End Sub

'@label:ImportSpecs
'@prop-title Import export specifications from another worksheet
'@param fromWksh Source worksheet containing the export table.
'@param fromStartRow First row of the source data.
'@param fromStartColumn First column of the source data.
'@throws ProjectError.ObjectNotInitialized when the source worksheet is missing.
Public Sub ImportSpecs(ByVal fromWksh As Worksheet, _
                       ByVal fromStartRow As Long, _
                       ByVal fromStartColumn As Long)

    EnsureDataReady
    If fromWksh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Source worksheet is missing"

    this.expData.Import fromWksh, fromStartRow, fromStartColumn
    InvalidateCaches
End Sub

'@label:ExportSpecs
'@prop-title Export specifications into another workbook
'@param toWkb Destination workbook.
'@param Hide Optional visibility flag for the exported sheet.
'@throws ProjectError.ObjectNotInitialized when the target workbook is missing.
Public Sub ExportSpecs(ByVal toWkb As Workbook, _
                       Optional ByVal Hide As Long = xlSheetHidden)

    EnsureDataReady
    If toWkb Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Target workbook is missing"

    this.expData.Export toWkb, Hide:=Hide
End Sub

'@label:InvalidateCaches
'@prop-title Reset cached column lookups and variables
Public Sub InvalidateCaches()
    ResetCaches
End Sub

'@section FileName
'===============================================================================
'@label:ExportFileName
'@prop-title Build the filename for a specific export definition
'@param exportNumber Target export definition number starting at 1.
'@param dict Dictionary providing metadata for chunk substitutions.
'@param pass Password provider used to append version/public key information.
'@param exportAll Optional flag forcing the "export all" naming schema.
'@return Sanitised filename incorporating dictionary-driven chunks and metadata.
'@throws ProjectError.InvalidArgument when the export number is invalid.
'@throws ProjectError.ObjectNotInitialized when required collaborators are missing.
Public Function ExportFileName(ByVal exportNumber As Long, _
                               ByVal dict As ILLdictionary, _
                               ByVal pass As IPasswords, _
                               Optional ByVal exportAll As Boolean = False) As String

    Dim currWb As Workbook
    Dim template As String
    Dim sanitizedPath As String
    Dim suffix As String

    EnsureDataReady
    EnsureWorksheetReady
    ValidatePasswords pass
    EnsureDictionaryReady dict

    Set currWb = this.exportSheet.Parent

    If exportAll Then
        ExportFileName = SanitizeChunk(WorkbookBaseName(currWb)) & "_export_all__" & BuildTimestamp
        Exit Function
    End If

    ValidateExportNumber exportNumber

    If Not IsActive(exportNumber) Then
        LogInfo "Export " & CStr(exportNumber) & " is inactive; generating filename for traceability"
    End If

    template = Trim$(ReadValue(COL_FILE_NAME, exportNumber, required:=True))

    If LenB(template) = 0 Then
        LogInfo "Empty file name template for export " & CStr(exportNumber)
        template = DefaultTemplate(exportNumber)
    End If

    sanitizedPath = ComposeTemplatePath(template, dict)
    suffix = BuildVersionSuffix(pass)

    ExportFileName = sanitizedPath & suffix & "__" & BuildTimestamp
End Function

'@section Utilities
'===============================================================================
'@label:IsActive
'@prop-title Determine whether an export is marked active
'@param exportNumber Export definition number to inspect.
'@return True when the status column contains "active" (case-insensitive).
Public Function IsActive(ByVal exportNumber As Long) As Boolean
    Dim statusValue As String
    statusValue = ReadValue(COL_STATUS, exportNumber, required:=False)
    IsActive = (StrComp(statusValue, STATUS_ACTIVE, vbTextCompare) = 0)
End Function

'@label:ColumnValue
'@prop-title Retrieve a column value for a given export row
'@param exportNumber Export definition number to read.
'@param columnName Column header to retrieve.
'@return Column value as string (empty string when not found).
Public Function ColumnValue(ByVal exportNumber As Long, ByVal columnName As String) As String
    ColumnValue = ReadValue(columnName, exportNumber, required:=False)
End Function

'@label:ActiveExportNumbers
'@prop-title Enumerate export numbers marked as active
'@return BetterArray containing active export numbers using 1-based numbering.
Public Function ActiveExportNumbers() As BetterArray
    Dim total As Long
    Dim idx As Long
    Dim result As BetterArray

    EnsureDataReady
    total = NumberOfExports()

    Set result = New BetterArray
    result.LowerBound = 1

    For idx = 1 To total
        If IsActive(idx) Then result.Push idx
    Next idx

    Set ActiveExportNumbers = result
End Function

'@section Private Helpers
'===============================================================================
'@label:EnsureDataReady
'@prop-title Ensure the datasheet state is initialised
Private Sub EnsureDataReady()
    If this.expData Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Export data store not initialised"
    End If
End Sub

'@label:EnsureWorksheetReady
'@prop-title Ensure the worksheet reference is initialised
Private Sub EnsureWorksheetReady()
    If this.exportSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Export worksheet not initialised"
    End If
End Sub

'@label:EnsureDictionaryReady
'@prop-title Validate dictionary dependency
Private Sub EnsureDictionaryReady(ByVal dict As ILLdictionary)
    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Dictionary is required for filename generation"
    End If
End Sub

'@label:ValidateExportNumber
'@prop-title Validate export number bounds
Private Sub ValidateExportNumber(ByVal exportNumber As Long)
    Dim total As Long
    total = NumberOfExports()

    If total = 0 Then
        ThrowError ProjectError.ElementNotFound, "No export specifications available"
    End If

    If exportNumber < 1 Or exportNumber > total Then
        ThrowError ProjectError.InvalidArgument, _
                  "Export number " & CStr(exportNumber) & " is out of range"
    End If
End Sub

'@label:ValidatePasswords
'@prop-title Validate password provider dependency
Private Sub ValidatePasswords(ByVal pass As IPasswords)
    If pass Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Passwords object is required"
    End If
End Sub

'@label:ExportTable
'@prop-title Retrieve export table wrapper
Private Function ExportTable() As ICustomTable
    Dim listObjects As ListObjects
    Dim lo As ListObject

    EnsureWorksheetReady

    Set listObjects = this.exportSheet.ListObjects
    If listObjects Is Nothing Or listObjects.Count = 0 Then
        ThrowError ProjectError.ElementNotFound, "No table found on export worksheet"
    End If

    Set lo = listObjects(1)
    Set ExportTable = CustomTable.Create(lo, COL_EXPORT_NUMBER)
End Function

'@label:ReadValue
'@prop-title Read a column value from the cached datasheet
Private Function ReadValue(ByVal columnName As String, _
                           ByVal exportNumber As Long, _
                           Optional ByVal required As Boolean = False) As String

    Dim columnRange As Range

    Set columnRange = ResolveColumnRange(columnName, required)

    If columnRange Is Nothing Then Exit Function

    If exportNumber < 1 Or exportNumber > columnRange.Rows.Count Then
        If required Then
            ThrowError ProjectError.InvalidArgument, _
                      "Export number " & CStr(exportNumber) & " is outside the specification range"
        Else
            Exit Function
        End If
    End If

    ReadValue = CStr(columnRange.Cells(exportNumber, 1).Value)
End Function

'@label:ResolveColumnRange
'@prop-title Resolve and cache a specification column
Private Function ResolveColumnRange(ByVal columnName As String, _
                                    ByVal required As Boolean) As Range

    Dim resolved As Range

    On Error GoTo ColumnMissing
        Set resolved = this.expData.DataRange(colName:=columnName, _
                                             includeHeaders:=False, _
                                             strictSearch:=True, _
                                             matchCase:=False)
    On Error GoTo 0

    Set ResolveColumnRange = resolved
    Exit Function

ColumnMissing:
    On Error GoTo 0
    If required Then
        ThrowError ProjectError.ElementNotFound, _
                   "Export specification column '" & columnName & "' not found"
    End If
End Function

'@label:ResetCaches
'@prop-title Reset cached references and diagnostics
Public Sub ResetCaches()
    Set this.variables = Nothing
    this.hasChecks = False
    this.checkCounter = 0
    Set this.checks = Nothing
    this.sanitizeList = Empty
End Sub

'@label:EnsureVariablesReady
'@prop-title Lazily instantiate dictionary-backed variable helper
Private Sub EnsureVariablesReady(ByVal dict As ILLdictionary)
    EnsureDictionaryReady dict
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(dict)
    End If
End Sub

'@label:ComposeTemplatePath
'@prop-title Build the filename path from template chunks
Private Function ComposeTemplatePath(ByVal template As String, _
                                     ByVal dict As ILLdictionary) As String

    Dim pathParts As BetterArray
    Dim chunks As Variant
    Dim idx As Long
    Dim resolved As String
    Dim workbook As Workbook

    EnsureVariablesReady dict
    Set workbook = this.exportSheet.Parent

    chunks = Split(template, "+")
    Set pathParts = New BetterArray
    pathParts.LowerBound = LBound(chunks)

    For idx = LBound(chunks) To UBound(chunks)
        resolved = ResolveTemplateChunk(CStr(chunks(idx)), workbook)
        pathParts.Push resolved
    Next idx

    ComposeTemplatePath = pathParts.ToString(separator:="__", _
                                            OpeningDelimiter:=vbNullString, _
                                            ClosingDelimiter:=vbNullString, _
                                            QuoteStrings:=False)
End Function

'@label:ResolveTemplateChunk
'@prop-title Resolve a single template chunk into a safe string
Private Function ResolveTemplateChunk(ByVal rawChunk As String, _
                                      ByVal workbook As Workbook) As String

    Dim chunk As String
    Dim value As String

    chunk = Trim$(rawChunk)
    If LenB(chunk) = 0 Then
        ResolveTemplateChunk = "chunk"
        Exit Function
    End If

    If IsLiteralTemplateChunk(chunk) Then
        ResolveTemplateChunk = SanitizeChunk(StripLiteralChunk(chunk))
        Exit Function
    End If

    value = ResolveVariableChunk(chunk, workbook)

    If LenB(value) = 0 Then
        LogInfo "Chunk resolved to empty string: " & chunk
        value = SanitizeChunk(chunk)
    End If

    ResolveTemplateChunk = value
End Function

'@label:IsLiteralTemplateChunk
'@prop-title Determine whether a chunk represents a literal string
Private Function IsLiteralTemplateChunk(ByVal chunk As String) As Boolean
    Dim firstChar As String
    Dim lastChar As String

    If Len(chunk) < 2 Then Exit Function

    firstChar = Left$(chunk, 1)
    lastChar = Right$(chunk, 1)

    If firstChar = Chr$(34) And lastChar = Chr$(34) Then
        IsLiteralTemplateChunk = True
        Exit Function
    End If

    If firstChar = "'" And lastChar = "'" Then
        IsLiteralTemplateChunk = True
        Exit Function
    End If
End Function

'@label:StripLiteralChunk
'@prop-title Remove wrapping quotes from literal chunk values
Private Function StripLiteralChunk(ByVal chunk As String) As String
    Dim lengthValue As Long

    lengthValue = Len(chunk)
    If lengthValue >= 2 Then
        If (Left$(chunk, 1) = Chr$(34) And Right$(chunk, 1) = Chr$(34)) _
           Or (Left$(chunk, 1) = "'" And Right$(chunk, 1) = "'") Then
            StripLiteralChunk = Mid$(chunk, 2, lengthValue - 2)
            Exit Function
        End If
    End If

    StripLiteralChunk = chunk
End Function

'@label:ResolveVariableChunk
'@prop-title Resolve dictionary-driven chunk value
Private Function ResolveVariableChunk(ByVal chunk As String, _
                                       ByVal workbook As Workbook) As String

    Dim vars As ILLVariables
    Dim sheetName As String
    Dim sheetType As String
    Dim targetSheet As Worksheet
    Dim namedRange As Range
    Dim resolved As String

    Set vars = this.variables

    If vars Is Nothing Then Exit Function
    If Not vars.Contains(chunk) Then Exit Function

    sheetName = vars.SheetName(chunk)
    sheetType = vars.Value("sheet type", chunk)

    If LenB(sheetName) = 0 Then Exit Function
    If StrComp(sheetType, "vlist1D", vbTextCompare) <> 0 Then Exit Function

    Set targetSheet = TryGetWorksheet(workbook, sheetName)
    If targetSheet Is Nothing Then
        LogInfo "Worksheet '" & sheetName & "' for chunk '" & chunk & "' is missing"
        Exit Function
    End If

    Set namedRange = TryGetNamedRange(workbook, chunk, targetSheet)

    If namedRange Is Nothing Then
        LogInfo "Named range for chunk '" & chunk & "' is missing", checkingWarning
        Exit Function
    End If

    resolved = CStr(namedRange.Value)
    resolved = SanitizeChunk(resolved)

    ResolveVariableChunk = resolved
End Function

'@label:TryGetWorksheet
'@prop-title Safely retrieve a worksheet by name
Private Function TryGetWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
        Set TryGetWorksheet = workbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

'@label:TryGetNamedRange
'@prop-title Safely retrieve a named range for a chunk
Private Function TryGetNamedRange(ByVal workbook As Workbook, _
                                  ByVal chunk As String, _
                                  ByVal targetSheet As Worksheet) As Range

    On Error Resume Next
        Set TryGetNamedRange = workbook.Names(chunk).RefersToRange
    On Error GoTo 0

    If TryGetNamedRange Is Nothing Then
        On Error Resume Next
            Set TryGetNamedRange = targetSheet.Range(chunk)
        On Error GoTo 0
    End If
End Function

'@label:SanitizeChunk
'@prop-title Sanitise individual chunk values for filenames
Private Function SanitizeChunk(ByVal chunk As String) As String
    Dim sanitized As String
    sanitized = Trim$(chunk)
    sanitized = ReplaceInvalidCharacters(sanitized)
    If LenB(sanitized) = 0 Then sanitized = "chunk"
    SanitizeChunk = sanitized
End Function

'@label:ReplaceInvalidCharacters
'@prop-title Replace characters unsupported in filenames
Private Function ReplaceInvalidCharacters(ByVal value As String) As String
    Dim invalidChar As Variant
    Dim newValue As String

    EnsureSanitizeList

    For Each invalidChar In this.sanitizeList
        newValue = Replace(value, CStr(invalidChar), "_")
    Next invalidChar

    ReplaceInvalidCharacters = newValue
End Function

'@label:EnsureSanitizeList
'@prop-title Initialise invalid filename character list
Private Sub EnsureSanitizeList()
    If IsEmpty(this.sanitizeList) Then
        this.sanitizeList = Array("<", ">", ":", "\", "/", "|", "?", "*", _
                                  Chr$(34), "&", ".", " ")
    End If
End Sub

'@label:BuildVersionSuffix
'@prop-title Compose version/public key suffix for filenames
Private Function BuildVersionSuffix(ByVal pass As IPasswords) As String
    Dim versionValue As String
    Dim publicKeyValue As String

    versionValue = SanitizeChunk(pass.Value("version"))
    publicKeyValue = SanitizeChunk(pass.Value("publickey"))

    BuildVersionSuffix = "__v" & versionValue & "-" & publicKeyValue
End Function

'@label:BuildTimestamp
'@prop-title Compose timestamp fragment for filenames
Private Function BuildTimestamp() As String
    BuildTimestamp = Format$(Now, TIMESTAMP_FORMAT)
End Function

'@label:DefaultTemplate
'@prop-title Provide fallback filename template
Private Function DefaultTemplate(ByVal exportNumber As Long) As String
    DefaultTemplate = "export" & CStr(exportNumber)
End Function


'@label:WorkbookBaseName
'@prop-title Remove Excel extensions from workbook name
Private Function WorkbookBaseName(ByVal workbook As Workbook) As String
    Dim cleaned As String

    cleaned = LCase$(workbook.Name)
    cleaned = Replace(cleaned, ".xlsm", vbNullString)
    cleaned = Replace(cleaned, ".xlsb", vbNullString)
    cleaned = Replace(cleaned, ".xlsx", vbNullString)
    cleaned = Replace(cleaned, ".xls", vbNullString)

    WorkbookBaseName = cleaned
End Function

'@label:LogInfo
'@prop-title Append informational diagnostics to the Checking log
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingInfo)
    If this.checks Is Nothing Then Set this.checks = Checking.Create(CHECKING_CONTEXT)
    this.checks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@section Error Handling
'===============================================================================
'@label:ThrowError
'@prop-title Raise a ProjectError coded runtime error
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'@section Interface Implementation
'===============================================================================
Private Property Get ILLExport_Data() As IDataSheet
    Set ILLExport_Data = Data
End Property

Private Property Get ILLExport_Worksheet() As Worksheet
    Set ILLExport_Worksheet = Worksheet
End Property

Private Property Get ILLExport_NumberOfExports() As Long
    ILLExport_NumberOfExports = NumberOfExports
End Property

Private Property Get ILLExport_HasCheckings() As Boolean
    ILLExport_HasCheckings = HasCheckings
End Property

Private Property Get ILLExport_CheckingValues() As IChecking
    Set ILLExport_CheckingValues = CheckingValues
End Property

Private Sub ILLExport_AddRows(Optional ByVal rowCount As Long = 1)
    AddRows rowCount
End Sub

Private Sub ILLExport_RemoveRows(Optional ByVal totalCount As Long = 0)
    RemoveRows totalCount
End Sub

Private Sub ILLExport_ImportSpecs(ByVal fromWksh As Worksheet, _
                                  ByVal fromStartRow As Long, _
                                  ByVal fromStartColumn As Long)
    ImportSpecs fromWksh, fromStartRow, fromStartColumn
End Sub

Private Sub ILLExport_ExportSpecs(ByVal toWkb As Workbook, _
                                  Optional ByVal Hide As Long = xlSheetHidden)
    ExportSpecs toWkb, Hide
End Sub

Private Sub ILLExport_InvalidateCaches()
    InvalidateCaches
End Sub

Private Function ILLExport_ExportFileName(ByVal exportNumber As Long, _
                                          ByVal dict As ILLdictionary, _
                                          ByVal pass As IPasswords, _
                                          Optional ByVal exportAll As Boolean = False) As String
    ILLExport_ExportFileName = ExportFileName(exportNumber, dict, pass, exportAll)
End Function

Private Function ILLExport_IsActive(ByVal exportNumber As Long) As Boolean
    ILLExport_IsActive = IsActive(exportNumber)
End Function

Private Function ILLExport_ColumnValue(ByVal exportNumber As Long, ByVal columnName As String) As String
    ILLExport_ColumnValue = ColumnValue(exportNumber, columnName)
End Function

Private Function ILLExport_ActiveExportNumbers() As BetterArray
    Set ILLExport_ActiveExportNumbers = ActiveExportNumbers
End Function
