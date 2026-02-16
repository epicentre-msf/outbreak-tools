VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist sheets of a dictionary"

'@PredeclaredId
'@Folder("Dictionary")
'@ModuleDescription("Linelist sheets of a dictionary")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing,  UseMeaningfulName, HungarianNotation, ExcelMemberMayReturnNothing
'@depends LLdictionary, LLVariables

Option Explicit

Implements ILLSheets

'@section Overview
'===============================================================================
'@details Coordinates read-only access to linelist sheet metadata stored in the
'dictionary. The class validates state on every call, provides rich
'error messages through ProjectError, and centralises frequently used
'lookups so callers avoid re-querying worksheet ranges.

'@section State
'===============================================================================

Private Type TLLSheets
    dict As ILLdictionary
    sheetRng As Range
    sheetTypeColumnIndex As Long
    tableNameColumnIndex As Long
    controlColumnIndex As Long
    variables As ILLVariables
End Type

Private this As TLLSheets

Private Const CLASSNAME As String = "LLSheets"
Private Const COL_SHEET_NAME As String = "sheet name"
Private Const COL_SHEET_TYPE As String = "sheet type"
Private Const COL_TABLE_NAME As String = "table name"
Private Const COL_CONTROL As String = "control"
Private Const COL_VARIABLE_NAME As String = "variable name"
Private Const SHEETTYPE_HORIZONTAL As String = "hlist2D"
Private Const SHEETTYPE_VERTICAL As String = "vlist1D"
Private Const SHEET_HORIZONTAL_DATA_ROWS As Long = 201
Private Const SHEET_START_ROW_HORIZONTAL As Long = 8
Private Const SHEET_START_ROW_VERTICAL As Long = 4
Private Const SHEET_START_COL_HORIZONTAL As Long = 1
Private Const SHEET_START_COL_VERTICAL As Long = 5

'@note Byte constants mirror SheetInfoType values while avoiding Enum usage, which keeps compatibility with macOS VBA.
Private Const LLSheetInfoFieldSheetType As Byte = SheetInfoType.SheetInfoSheetType
Private Const LLSheetInfoFieldTableName As Byte = SheetInfoType.SheetInfoSheetTable

Private Const LLSheetBoundTop As Byte = SheetBound.RowSart
Private Const LLSheetBoundBottom As Byte = SheetBound.RowEnd
Private Const LLSheetBoundLeft As Byte = SheetBound.ColStart
Private Const LLSheetBoundRight As Byte = SheetBound.ColEnd

'@section Factory
'===============================================================================

'@label:Create
'@sub-title Instantiate a sheet helper bound to the supplied dictionary
'@param dict ILLdictionary providing sheet metadata.
'@return ILLSheets fully initialised helper.
'@export
Public Function Create(ByVal dict As ILLdictionary) As ILLSheets

    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Unable to create LLSheets: dictionary reference is missing"
    End If

    With New LLSheets
        Set .Dictionary = dict
        Set Create = .Self
    End With
End Function

'@section Public Accessors
'===============================================================================

'@label:Self
'@sub-title Expose the interface reference of this predeclared instance
'@return ILLSheets reference to the current object.
Public Property Get Self() As ILLSheets
    Set Self = Me
End Property

'@label:Dictionary
'@sub-title Retrieve the backing dictionary
'@return ILLdictionary reference captured during initialisation.
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

'@sub-title Store the backing dictionary and prime cached lookups
'@param dict ILLdictionary holding linelist metadata.
Public Property Set Dictionary(ByVal dict As ILLdictionary)
    ValidateDictionary dict
    Set this.dict = dict
    ResetCaches
    Set this.sheetRng = dict.DataRange(colName:=COL_SHEET_NAME, includeHeaders:=True)
End Property

'@label:SheetRange
'@sub-title Return the cached range referencing sheet names
'@return Range spanning the sheet-name column including the header.
Private Property Get SheetRange(Optional ByVal requirePrepared As Boolean = False) As Range
    EnsureDictionaryReady requirePrepared
    Set SheetRange = this.sheetRng
End Property

'@section Sheet Queries
'===============================================================================

'@label:Contains
'@sub-title Determine whether the dictionary includes the supplied sheet
'@param sheetName String. Sheet identifier to locate.
'@return True when the sheet exists.
Private Function Contains(ByVal sheetName As String) As Boolean
    Dim found As Range

    If LenB(sheetName) = 0 Then Exit Function
    If StrComp(sheetName, COL_SHEET_NAME, vbTextCompare) = 0 Then Exit Function

    Set found = FindSheetCell(sheetName)
    Contains = Not (found Is Nothing)
End Function

'@label:RowIndex
'@sub-title Lookup the worksheet row hosting the sheet metadata
'@param sheetName String. Sheet identifier to locate.
'@return Worksheet row index for the sheet entry, or 0 when not found.
Private Function RowIndex(ByVal sheetName As String) As Long
    Dim found As Range

    If LenB(sheetName) = 0 Then Exit Function

    Set found = FindSheetCell(sheetName)
    If Not (found Is Nothing) Then RowIndex = found.Row
End Function

'@label:SheetInfo
'@sub-title Fetch additional metadata about a sheet
'@param sheetName String. Sheet identifier.
'@param info Byte. Requested metadata selector.
'@return Requested metadata value; raises when selector is unknown.
Public Function SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = SheetInfoType.SheetInfoSheetType) As String
    sheetInfo = SheetInfoValue(sheetName, info)
End Function

'@label:VarList
'@sub-title Retrieve the ordered list of variables belonging to a sheet
'@param sheetName String. Sheet identifier.
'@return BetterArray containing variable names for the sheet.
Private Function Varlist(ByVal sheetName As String) As BetterArray
    Dim listData As BetterArray

    EnsureSheetExists sheetName
    Set listData = this.dict.Data.FilterData(COL_SHEET_NAME, sheetName, COL_VARIABLE_NAME, matchCase:=False)
    Set Varlist = listData.Clone()
End Function

'@label:NumberVars
'@sub-title Count the variables associated with a sheet
'@param sheetName String. Sheet identifier.
'@return Long number of variables present on the sheet.
Private Function NumberVars(ByVal sheetName As String) As Long
    Dim vars As BetterArray

    Set vars = Varlist(sheetName)
    NumberVars = vars.Length
End Function

'@label:DataBounds
'@sub-title Compute the data boundaries for a sheet layout
'@param sheetName String. Sheet identifier.
'@param Bound Byte. Selector from SheetBound enumeration.
'@return Long representing the requested bound.
Private Function DataBounds(ByVal sheetName As String, _
                            Optional ByVal Bound As Byte = SheetBound.RowSart) As Long
    Dim sheetType As String
    Dim varCount As Long
    Dim offset As Long

    sheetType = SheetInfoValue(sheetName, LLSheetInfoFieldSheetType)

    Select Case Bound
        Case LLSheetBoundTop
            Select Case sheetType
                Case SHEETTYPE_HORIZONTAL
                    DataBounds = SHEET_START_ROW_HORIZONTAL
                Case SHEETTYPE_VERTICAL
                    DataBounds = SHEET_START_ROW_VERTICAL
                Case Else
                    ThrowError ProjectError.InvalidArgument, _
                               "Unknown sheet type '" & sheetType & "' for bound lookup"
            End Select

        Case LLSheetBoundBottom
            Select Case sheetType
                Case SHEETTYPE_HORIZONTAL
                    DataBounds = SHEET_START_ROW_HORIZONTAL + SHEET_HORIZONTAL_DATA_ROWS
                Case SHEETTYPE_VERTICAL
                    varCount = NumberVars(sheetName)
                    offset = IIf(varCount > 0, varCount - 1, 0)
                    DataBounds = SHEET_START_ROW_VERTICAL + offset
                Case Else
                    ThrowError ProjectError.InvalidArgument, _
                               "Unknown sheet type '" & sheetType & "' for bound lookup"
            End Select

        Case LLSheetBoundLeft
            Select Case sheetType
                Case SHEETTYPE_HORIZONTAL
                    DataBounds = SHEET_START_COL_HORIZONTAL
                Case SHEETTYPE_VERTICAL
                    DataBounds = SHEET_START_COL_VERTICAL
                Case Else
                    ThrowError ProjectError.InvalidArgument, _
                               "Unknown sheet type '" & sheetType & "' for bound lookup"
            End Select

        Case LLSheetBoundRight
            Select Case sheetType
                Case SHEETTYPE_HORIZONTAL
                    varCount = NumberVars(sheetName)
                    offset = IIf(varCount > 0, varCount - 1, 0)
                    DataBounds = SHEET_START_COL_HORIZONTAL + offset
                Case SHEETTYPE_VERTICAL
                    DataBounds = SHEET_START_COL_VERTICAL
                Case Else
                    ThrowError ProjectError.InvalidArgument, _
                               "Unknown sheet type '" & sheetType & "' for bound lookup"
            End Select

        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unsupported bound selector: " & CStr(Bound)
    End Select
End Function

'@label:ContainsControl
'@sub-title Test whether a sheet exposes a control of the requested type
'@param sheetName String. Sheet identifier.
'@param Control String. Control type to search for.
'@param colName String. Column hosting the control types. Defaults to "control".
'@return True when the control type exists on the sheet.
Private Function ContainsControl(ByVal sheetName As String, _
                                 Optional ByVal Control As String = "list_auto", _
                                 Optional ByVal colName As String = COL_CONTROL) As Boolean
    Dim values As BetterArray

    EnsureSheetExists sheetName
    EnsureControlColumnExists colName

    Set values = this.dict.Data.FilterData(COL_SHEET_NAME, sheetName, colName, matchCase:=False)
    ContainsControl = values.Includes(Control)
End Function

'@label:VariableAddress
'@sub-title Build an Excel address pointing to the variable slot on its sheet
'@param varName String. Variable identifier.
'@param onSheet Optional String. Sheet where the reference will be consumed; suppresses sheet prefix when matching.
'@return String Excel address for the variable.
Private Function VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String
    Dim vars As ILLVariables
    Dim targetSheet As String
    Dim sheetType As String
    Dim sheetRow As Long
    Dim sheetCol As Long
    Dim varIndex As Long
    Dim wksh As Worksheet
    Dim prefix As String

    EnsureDictionaryReady True

    If LenB(varName) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "Variable name cannot be empty when requesting an address"
    End If

    Set vars = VariablesProvider()

    If Not vars.Contains(varName) Then
        ThrowError ProjectError.ElementNotFound, _
                   "Variable '" & varName & "' is not present in the dictionary"
    End If

    targetSheet = vars.Value(COL_SHEET_NAME, varName)
    If LenB(targetSheet) = 0 Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "Variable '" & varName & "' has no sheet assignment"
    End If

    varIndex = vars.Index(varName)
    If varIndex <= 0 Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "Dictionary is missing the column index for variable '" & varName & "'"
    End If

    sheetType = SheetInfoValue(targetSheet, LLSheetInfoFieldSheetType)
    Set wksh = this.dict.Data.Wksh

    If LenB(onSheet) = 0 Or StrComp(onSheet, targetSheet, vbTextCompare) <> 0 Then
        prefix = BuildSheetPrefix(targetSheet)
    End If

    Select Case sheetType
        Case SHEETTYPE_HORIZONTAL
            sheetRow = DataBounds(targetSheet, LLSheetBoundTop)
            VariableAddress = prefix & wksh.Cells(sheetRow + 1, varIndex).Address(RowAbsolute:=False, ColumnAbsolute:=False)

        Case SHEETTYPE_VERTICAL
            sheetCol = DataBounds(targetSheet, LLSheetBoundLeft)
            VariableAddress = prefix & wksh.Cells(varIndex, sheetCol).Address(RowAbsolute:=True, ColumnAbsolute:=True)

        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unknown sheet type '" & sheetType & "' while building variable address"
    End Select
End Function

'@section State Validation
'===============================================================================

'@label:EnsureDictionaryReady
'@sub-title Confirm the backing dictionary and cached range are initialised
'@param requirePrepared Optional Boolean. When True, the dictionary must be prepared.
Private Sub EnsureDictionaryReady(Optional ByVal requirePrepared As Boolean = False)
    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "LLSheets has not been initialised. Call Create before use"
    End If

    If requirePrepared And Not this.dict.Prepared Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary must be prepared before requesting variable-dependent data"
    End If

    If this.sheetRng Is Nothing Then
        Set this.sheetRng = this.dict.DataRange(colName:=COL_SHEET_NAME, includeHeaders:=True)
    End If
End Sub

'@label:EnsureSheetExists
'@sub-title Raise a descriptive error when the sheet does not exist
'@param sheetName String. Sheet identifier to validate.
Private Sub EnsureSheetExists(ByVal sheetName As String)
    Dim rowNum As Long

    rowNum = EnsureSheetRow(sheetName)
    If rowNum = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "Sheet '" & sheetName & "' is not defined in the dictionary"
    End If
End Sub

'@label:EnsureSheetRow
'@sub-title Retrieve the worksheet row for the sheet or raise when absent
'@param sheetName String. Sheet identifier.
'@return Long worksheet row index.
Private Function EnsureSheetRow(ByVal sheetName As String) As Long
    Dim found As Range

    EnsureValidSheetName sheetName
    Set found = FindSheetCell(sheetName)

    If found Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Sheet '" & sheetName & "' is not defined in the dictionary"
    Else
        EnsureSheetRow = found.Row
    End If
End Function

'@label:EnsureControlColumnExists
'@sub-title Guarantee the control column exists before filtering
'@param colName String. Column header to validate.
Private Sub EnsureControlColumnExists(ByVal colName As String)
    Dim targetName As String

    targetName = IIf(LenB(colName) = 0, COL_CONTROL, colName)

    If StrComp(targetName, COL_CONTROL, vbTextCompare) = 0 Then
        If this.controlColumnIndex = 0 Then
            this.controlColumnIndex = ColumnIndex(targetName, allowMissing:=False)
        End If
    Else
        Call ColumnIndex(targetName, allowMissing:=False)
    End If
End Sub

'@label:EnsureValidSheetName
'@sub-title Validate that the supplied sheet name is usable
'@param sheetName String. Sheet identifier to validate.
Private Sub EnsureValidSheetName(ByVal sheetName As String)
    If LenB(sheetName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheet name cannot be empty"
    End If

    If StrComp(sheetName, COL_SHEET_NAME, vbTextCompare) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "Sheet name cannot match the dictionary header value"
    End If
End Sub

'@section Lookup Helpers
'===============================================================================

'@label:SheetInfoValue
'@sub-title Resolve sheet metadata based on the selector
'@param sheetName String. Sheet identifier.
'@param infoSelector Byte. Selector from LLSheetInfoField.
'@return String metadata value.
Private Function SheetInfoValue(ByVal sheetName As String, ByVal infoSelector As Byte) As String
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim value As Variant

    rowIndex = EnsureSheetRow(sheetName)

    Select Case infoSelector
        Case LLSheetInfoFieldSheetType
            columnIndex = SheetTypeColumnIndex()
        Case LLSheetInfoFieldTableName
            columnIndex = TableNameColumnIndex()
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unknown sheet info selector: " & CStr(infoSelector)
    End Select

    value = this.dict.Data.Wksh.Cells(rowIndex, columnIndex).Value
    SheetInfoValue = CStr(value)
End Function

'@label:SheetTypeColumnIndex
'@sub-title Cache the column index for sheet type lookups
'@return Long worksheet column index.
Private Function SheetTypeColumnIndex() As Long
    If this.sheetTypeColumnIndex = 0 Then
        this.sheetTypeColumnIndex = ColumnIndex(COL_SHEET_TYPE, allowMissing:=False)
    End If
    SheetTypeColumnIndex = this.sheetTypeColumnIndex
End Function

'@label:TableNameColumnIndex
'@sub-title Cache the column index for table name lookups and raise if missing
'@return Long worksheet column index.
Private Function TableNameColumnIndex() As Long
    If this.tableNameColumnIndex = 0 Then
        If Not this.dict.ColumnExists(COL_TABLE_NAME) Then
            ThrowError ProjectError.ElementNotFound, _
                       "Dictionary column '" & COL_TABLE_NAME & "' is required for this operation"
        End If
        this.tableNameColumnIndex = ColumnIndex(COL_TABLE_NAME, allowMissing:=False)
    End If
    TableNameColumnIndex = this.tableNameColumnIndex
End Function

'@label:VariablesProvider
'@sub-title Lazily initialise the LLVariables helper for address resolution
'@return ILLVariables cached instance.
Private Function VariablesProvider() As ILLVariables
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(this.dict)
    End If
    Set VariablesProvider = this.variables
End Function

'@label:ColumnIndex
'@sub-title Wrapper around IDataSheet.ColumnIndex with optional caching
'@param colName String. Column header to locate.
'@param allowMissing Boolean. When False, raises when the column is absent.
'@return Long worksheet column index.
Private Function ColumnIndex(ByVal colName As String, ByVal allowMissing As Boolean) As Long
    Dim idx As Long

    idx = this.dict.Data.ColumnIndex(colName, shouldExist:=Not allowMissing, matchCase:=False)
    If idx <= 0 And Not allowMissing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Dictionary column '" & colName & "' is required for this operation"
    End If

    ColumnIndex = idx
End Function

'@label:FindSheetCell
'@sub-title Locate the range corresponding to a sheet name
'@param sheetName String. Sheet identifier (may be empty).
'@return Range when found; Nothing otherwise.
Private Function FindSheetCell(ByVal sheetName As String) As Range
    Dim searchValue As String
    Dim result As Range

    EnsureDictionaryReady

    If LenB(sheetName) = 0 Then Exit Function

    searchValue = SanitizeSearchTerm(sheetName)
    On Error Resume Next
    Set result = this.sheetRng.Find(What:=searchValue, LookAt:=xlWhole, MatchCase:=True)
    On Error GoTo 0

    Set FindSheetCell = result
End Function

'@label:SanitizeSearchTerm
'@sub-title Escape Excel wildcard characters before using Range.Find
'@param value String. Raw search term.
'@return String sanitized search term.
Private Function SanitizeSearchTerm(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Replace(value, "~", "~~")
    cleaned = Replace(cleaned, "*", "~*")
    cleaned = Replace(cleaned, "?", "~?")
    cleaned = Replace(cleaned, "[", "~[")
    cleaned = Replace(cleaned, "]", "~]")
    cleaned = Replace(cleaned, "#", "~#")

    SanitizeSearchTerm = cleaned
End Function

'@label:BuildSheetPrefix
'@sub-title Produce a quoted sheet prefix for addresses
'@param sheetName String. Sheet identifier to encode.
'@return String sheet prefix including trailing exclamation mark.
Private Function BuildSheetPrefix(ByVal sheetName As String) As String
    BuildSheetPrefix = "'" & Replace(sheetName, "'", "''") & "'!"
End Function

'@label:ResetCaches
'@sub-title Reset cached lookups following dictionary changes
Private Sub ResetCaches()
    this.sheetTypeColumnIndex = 0
    this.tableNameColumnIndex = 0
    this.controlColumnIndex = 0
    Set this.variables = Nothing
    Set this.sheetRng = Nothing
End Sub

'@label:InvalidateCaches
Public Sub InvalidateCaches()
    ResetCaches
End Sub

'@label:ValidateDictionary
'@sub-title Ensure the supplied dictionary exposes the required columns
'@param dict ILLdictionary to validate.
Private Sub ValidateDictionary(ByVal dict As ILLdictionary)
    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Cannot assign a null dictionary to LLSheets"
    End If

    If dict.Data Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary DataSheet is not initialised"
    End If

    EnsureColumnExists dict, COL_SHEET_NAME
    EnsureColumnExists dict, COL_SHEET_TYPE
    EnsureColumnExists dict, COL_VARIABLE_NAME
End Sub

'@label:EnsureColumnExists
'@sub-title Validate that a dictionary column is present
'@param dict ILLdictionary to inspect.
'@param colName String. Column header to verify.
Private Sub EnsureColumnExists(ByVal dict As ILLdictionary, ByVal colName As String)
    If Not dict.ColumnExists(colName) Then
        ThrowError ProjectError.ElementNotFound, _
                   "Dictionary column '" & colName & "' is required for LLSheets"
    End If
End Sub

'@section Error Handling
'===============================================================================

'@label:ThrowError
'@sub-title Raise a ProjectError-coded runtime error
'@param errNumb ProjectError enumeration value.
'@param errorMessage String human readable description.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'@section Interface
'===============================================================================

Private Property Get ILLSheets_Dictionary() As ILLdictionary
    Set ILLSheets_Dictionary = Dictionary()
End Property

Private Function ILLSheets_SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = SheetInfoType.SheetInfoSheetType) As String
    ILLSheets_SheetInfo = sheetInfo(sheetName, info)
End Function

Private Function ILLSheets_Contains(ByVal sheetName As String) As Boolean
    ILLSheets_Contains = Contains(sheetName)
End Function

Private Function ILLSheets_DataBounds(ByVal sheetName As String, _
                                     Optional ByVal Bound As Byte = SheetBound.RowSart) As Long
    ILLSheets_DataBounds = DataBounds(sheetName, Bound)
End Function

Private Function ILLSheets_ContainsControl(ByVal sheetName As String, _
                                          Optional ByVal Control As String = "list_auto", _
                                          Optional ByVal colName As String = COL_CONTROL) As Boolean
    ILLSheets_ContainsControl = ContainsControl(sheetName, Control, colName)
End Function

Private Function ILLSheets_VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String
    ILLSheets_VariableAddress = VariableAddress(varName, onSheet)
End Function

Private Function ILLSheets_RowIndex(ByVal sheetName As String) As Long
    ILLSheets_RowIndex = RowIndex(sheetName)
End Function

Private Function ILLSheets_NumberOfVars(ByVal sheetName As String) As Long
    ILLSheets_NumberOfVars = NumberVars(sheetName)
End Function

Private Sub ILLSheets_InvalidateCaches()
    InvalidateCaches
End Sub
