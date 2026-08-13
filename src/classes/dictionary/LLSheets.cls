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
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================

'@description Bound selectors used by DataBounds. The original spelling of
'@description RowSart is preserved for backwards compatibility.
Public Enum SheetBound
    RowSart = 1
    RowEnd = 2
    ColStart = 3
    ColEnd = 4
End Enum

'@description Metadata selectors consumed by SheetInfo.
Public Enum SheetInfoType
    SheetInfoSheetType = 1
    SheetInfoSheetTable = 2
End Enum

'@class LLSheets
'@description
'Answers questions about ONE sheet of a linelist: the dictionary row that
'describes it, its type, its table name, how many variables it holds, and where
'a variable sits on it. Contains says whether the dictionary knows the sheet,
'RowIndex gives its dictionary row, SheetInfo answers its type or its table
'name, NumberOfVars counts its variables, DataBounds answers the first and the
'last row and column of its data, ContainsControl says whether a control
'appears on it, and VariableAddress gives the address of one variable. Every
'read checks the state it needs and raises a ProjectError whose message names
'the sheet.
'
'WHAT THE INSTANCE HOLDS
'-------------------------------------------------------------------------------
'The dictionary row is held per sheet name. Before a held row is handed out,
'the cell it points at is read once and compared with the name it should carry,
'so a dictionary whose rows moved gives a fresh row.
'
'The variable count is held per sheet name until InvalidateCaches.
'LLdictionary.Prepare inserts geo rows, so an object built before preparation
'answers over the dictionary as it stood then.
'
'The sheet-type and table-name column indexes are held too.
'
'The preparation flag is asked of the dictionary on every call that needs it.
'LLdictionary.Prepared holds a True answer itself and drops it in every routine
'that can turn True back into False. This class has no way to see a Clean or
'an AddColumn run on the dictionary, so it keeps no copy of that flag.
'@depends BetterArray, DataSheet, LLdictionary, LLVariables

'@section State
'===============================================================================

Private Type TLLSheets
    dict As LLdictionary
    hostWksh As Worksheet
    sheetRng As Range
    sheetTypeColumnIndex As Long
    tableNameColumnIndex As Long
    variables As LLVariables
    cachedNames() As String
    cachedRows() As Long
    cachedCounts() As Long
    cachedFarIndexes() As Long
    cachedSlots As Long
    sealed As Boolean
End Type

Private this As TLLSheets

Private Const CLASSNAME As String = "LLSheets"
Private Const COL_SHEET_NAME As String = "sheet name"
Private Const COL_SHEET_TYPE As String = "sheet type"
Private Const COL_TABLE_NAME As String = "table name"
Private Const COL_CONTROL As String = "control"
Private Const COL_VARIABLE_NAME As String = "variable name"
Private Const COL_COLUMN_INDEX As String = "column index"
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
'@param dict LLdictionary providing sheet metadata.
'@return LLSheets fully initialised helper.
'@throws ProjectError.ObjectNotInitialized When dict is Nothing or has no DataSheet.
'@throws ProjectError.ElementNotFound When a required dictionary column is missing.
'@export
Public Function Create(ByVal dict As LLdictionary) As LLSheets

    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Unable to create LLSheets: dictionary reference is missing"
    End If

    Dim newSheets As LLSheets
    Set newSheets = New LLSheets
    Set newSheets.Dictionary = dict
    newSheets.Seal
    Set Create = newSheets
End Function

'@section Public Accessors
'===============================================================================

'@label:Dictionary
'@sub-title Retrieve the backing dictionary
'@return LLdictionary reference captured during initialisation.
Public Property Get Dictionary() As LLdictionary
    Set Dictionary = this.dict
End Property

'@sub-title Store the backing dictionary and prime cached lookups
'@param dict LLdictionary holding linelist metadata.
Public Property Set Dictionary(ByVal dict As LLdictionary)
    GuardNotSealed "Dictionary"
    ValidateDictionary dict
    Set this.dict = dict
    ResetCaches
    Set this.sheetRng = dict.DataRange(colName:=COL_SHEET_NAME, includeHeaders:=True)
End Property

'@label:Seal
'@sub-title Prevent further changes to creation-only setters
Public Sub Seal()
    this.sealed = True
End Sub

'@label:GuardNotSealed
'@sub-title Guard creation-only setters after sealing
'@param propName String. Name used in the message.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@section Sheet Queries
'===============================================================================

'@label:Contains
'@sub-title Determine whether the dictionary includes the supplied sheet
'@details
'The match is case-insensitive, the same rule Excel applies to worksheet names.
'A sheet stored as "Data" and asked for as "data" is the same sheet, and
'`LLDataEntry` refuses to build a sheet this call reports as missing.
'@param sheetName String. Sheet identifier to locate.
'@return True when the sheet exists.
'@export
Public Function Contains(ByVal sheetName As String) As Boolean
    If LenB(sheetName) = 0 Then Exit Function
    If StrComp(sheetName, COL_SHEET_NAME, vbTextCompare) = 0 Then Exit Function

    Contains = (ResolveSheetRow(sheetName) > 0)
End Function

'@label:RowIndex
'@sub-title Lookup the worksheet row hosting the sheet metadata
'@details
'Answers 0 for a sheet the dictionary does not hold, and for the header row.
'The sheet-name range is captured with its header, so this used to hand back
'the header row for a call asking about the header text, and `LLDataEntry`
'places rows from the answer.
'@param sheetName String. Sheet identifier to locate.
'@return Worksheet row index for the sheet entry, or 0 when not found.
'@export
Public Function RowIndex(ByVal sheetName As String) As Long
    If LenB(sheetName) = 0 Then Exit Function

    RowIndex = ResolveSheetRow(sheetName)
End Function

'@label:SheetInfo
'@sub-title Fetch additional metadata about a sheet
'@param sheetName String. Sheet identifier.
'@param info Byte. Requested metadata selector, from SheetInfoType.
'@return Requested metadata value; raises when the selector is unknown.
'@throws ProjectError.ElementNotFound When the sheet or the metadata column is missing.
'@throws ProjectError.InvalidArgument When the selector is unknown.
'@export
Public Function SheetInfo(ByVal sheetName As String, Optional ByVal info As Byte = SheetInfoType.SheetInfoSheetType) As String
    SheetInfo = SheetInfoValue(sheetName, info)
End Function

'@label:NumberOfVars
'@sub-title Count the variables associated with a sheet
'@details
'The count is read once per sheet name and held. One AutoFilter pass answers
'it, and `DataBounds` asks for it twice per sheet on the `LLDataEntry` path,
'which asks again on its own.
'@param sheetName String. Sheet identifier.
'@return Long number of variables present on the sheet.
'@throws ProjectError.ElementNotFound When the sheet is missing from the dictionary.
'@export
Public Function NumberOfVars(ByVal sheetName As String) As Long
    Dim slot As Long

    EnsureSheetExists sheetName

    slot = CacheSlot(sheetName)
    If this.cachedCounts(slot) < 0 Then
        this.cachedCounts(slot) = CountVariables(sheetName)
    End If

    NumberOfVars = this.cachedCounts(slot)
End Function

'@label:count-variables
'@sub-title Count the dictionary rows that belong to one sheet
'@param sheetName String. Sheet identifier.
'@return Long number of variables on that sheet.
Private Function CountVariables(ByVal sheetName As String) As Long
    Dim listData As BetterArray

    Set listData = this.dict.Data.FilterData(COL_SHEET_NAME, sheetName, COL_VARIABLE_NAME, matchCase:=False)
    CountVariables = listData.Length
End Function

'@label:DataBounds
'@sub-title Compute the data boundaries for a sheet layout
'@details
'The four bounds of a horizontal sheet: top row 8, bottom row 8 + 201, left
'column 1, right column 1 + variables - 1. The four bounds of a vertical
'sheet: top row 4, bottom row 4 + variables - 1, left column 5, right
'column 5.
'@param sheetName String. Sheet identifier.
'@param Bound Byte. Selector from the SheetBound enumeration.
'@return Long representing the requested bound.
'@throws ProjectError.InvalidArgument When the sheet type or the selector is unknown.
'@export
Public Function DataBounds(ByVal sheetName As String, _
                            Optional ByVal Bound As Byte = SheetBound.RowSart) As Long
    DataBounds = BoundValue(sheetName, SheetInfoValue(sheetName, LLSheetInfoFieldSheetType), Bound)
End Function

'@label:bound-value
'@sub-title Compute one bound from a sheet type that is already resolved
'@details
'The sheet type is validated once, here. `VariableAddress` resolves the type
'for its own use and hands it down, so the address path reads the type cell
'once where it used to read it twice.
'@param sheetName String. Sheet identifier, used for the variable count.
'@param sheetType String. Sheet type already read from the dictionary.
'@param Bound Byte. Selector from the SheetBound enumeration.
'@return Long representing the requested bound.
'@throws ProjectError.InvalidArgument When the sheet type or the selector is unknown.
Private Function BoundValue(ByVal sheetName As String, _
                            ByVal sheetType As String, _
                            ByVal Bound As Byte) As Long
    Dim horizontal As Boolean
    Dim varCount As Long
    Dim offset As Long

    Select Case sheetType
        Case SHEETTYPE_HORIZONTAL
            horizontal = True
        Case SHEETTYPE_VERTICAL
            horizontal = False
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unknown sheet type '" & sheetType & "' for bound lookup"
    End Select

    Select Case Bound
        Case LLSheetBoundTop
            If horizontal Then
                BoundValue = SHEET_START_ROW_HORIZONTAL
            Else
                BoundValue = SHEET_START_ROW_VERTICAL
            End If

        Case LLSheetBoundLeft
            If horizontal Then
                BoundValue = SHEET_START_COL_HORIZONTAL
            Else
                BoundValue = SHEET_START_COL_VERTICAL
            End If

        Case LLSheetBoundBottom
            If horizontal Then
                BoundValue = SHEET_START_ROW_HORIZONTAL + SHEET_HORIZONTAL_DATA_ROWS
            Else
                varCount = NumberOfVars(sheetName)
                offset = IIf(varCount > 0, varCount - 1, 0)
                BoundValue = SHEET_START_ROW_VERTICAL + offset
                BoundValue = FarthestOf(BoundValue, sheetName)
            End If

        Case LLSheetBoundRight
            If horizontal Then
                varCount = NumberOfVars(sheetName)
                offset = IIf(varCount > 0, varCount - 1, 0)
                BoundValue = SHEET_START_COL_HORIZONTAL + offset
                BoundValue = FarthestOf(BoundValue, sheetName)
            Else
                BoundValue = SHEET_START_COL_VERTICAL
            End If

        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unsupported bound selector: " & CStr(Bound)
    End Select
End Function

'@label:ContainsControl
'@sub-title Test whether a sheet exposes a control of the requested type
'@details
'The default control value is "list_auto". `LLDataEntry` asks for "geo1",
'which the dictionary holds only after `LLdictionary.AppendGeoLines` has run.
'@param sheetName String. Sheet identifier.
'@param Control String. Control type to search for.
'@param colName String. Column hosting the control types. Defaults to "control".
'@return True when the control type exists on the sheet.
'@throws ProjectError.ElementNotFound When the sheet or the control column is missing.
'@export
Public Function ContainsControl(ByVal sheetName As String, _
                                 Optional ByVal ribbonControl As String = "list_auto", _
                                 Optional ByVal colName As String = COL_CONTROL) As Boolean
    Dim values As BetterArray

    EnsureSheetExists sheetName
    EnsureControlColumnExists colName

    Set values = this.dict.Data.FilterData(COL_SHEET_NAME, sheetName, colName, matchCase:=False)
    ContainsControl = values.Includes(ribbonControl)
End Function

'@label:VariableAddress
'@sub-title Build an Excel address pointing to the variable slot on its sheet
'@details
'The address is built from the DICTIONARY worksheet on purpose. `.Address`
'returns an A1 string that carries no sheet name, so only the row and the
'column number reach the result, and the sheet prefix is added here.
'
'A horizontal sheet gets a relative address and a vertical sheet gets an
'absolute one. `Formulas` and `VarWriter` drop both straight into formulas, so
'the two forms behave differently when Excel fills them down; that is the
'shape those two callers expect.
'@param varName String. Variable identifier.
'@param onSheet Optional String. Sheet where the reference will be consumed; suppresses sheet prefix when matching.
'@return String Excel address for the variable.
'@throws ProjectError.ObjectNotInitialized When the dictionary is missing or unprepared.
'@throws ProjectError.InvalidArgument When varName is empty or the sheet type is unknown.
'@throws ProjectError.ElementNotFound When the variable is absent from the dictionary.
'@export
Public Function VariableAddress(ByVal varName As String, Optional ByVal onSheet As String = vbNullString) As String
    Dim vars As LLVariables
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
    Set wksh = HostSheet()

    If LenB(onSheet) = 0 Or StrComp(onSheet, targetSheet, vbTextCompare) <> 0 Then
        prefix = BuildSheetPrefix(targetSheet)
    End If

    Select Case sheetType
        Case SHEETTYPE_HORIZONTAL
            sheetRow = BoundValue(targetSheet, sheetType, LLSheetBoundTop)
            VariableAddress = prefix & wksh.Cells(sheetRow + 1, varIndex).Address(RowAbsolute:=False, ColumnAbsolute:=False)

        Case SHEETTYPE_VERTICAL
            sheetCol = BoundValue(targetSheet, sheetType, LLSheetBoundLeft)
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
'@details
'`LLdictionary.Prepared` is asked on every call that requires preparation. It
'holds a True answer on the dictionary itself, so the read costs one Boolean
'once the dictionary is prepared, and it stays honest when another object
'undoes the preparation.
'@param requirePrepared Optional Boolean. When True, the dictionary must be prepared.
'@throws ProjectError.ObjectNotInitialized When the dictionary is missing or unprepared.
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
'@details
'`EnsureSheetRow` raises for a sheet the dictionary does not hold, so this is
'the same check under the name the two callers read better with.
'@param sheetName String. Sheet identifier to validate.
'@throws ProjectError.ElementNotFound When the sheet is missing from the dictionary.
Private Sub EnsureSheetExists(ByVal sheetName As String)
    EnsureSheetRow sheetName
End Sub

'@label:EnsureSheetRow
'@sub-title Retrieve the worksheet row for the sheet or raise when absent
'@param sheetName String. Sheet identifier.
'@return Long worksheet row index.
'@throws ProjectError.InvalidArgument When the name is empty or is the header value.
'@throws ProjectError.ElementNotFound When the sheet is missing from the dictionary.
Private Function EnsureSheetRow(ByVal sheetName As String) As Long
    Dim rowNum As Long

    EnsureValidSheetName sheetName
    rowNum = ResolveSheetRow(sheetName)

    If rowNum = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "Sheet '" & sheetName & "' is not defined in the dictionary"
    End If

    EnsureSheetRow = rowNum
End Function

'@label:EnsureControlColumnExists
'@sub-title Guarantee the control column exists before filtering
'@param colName String. Column header to validate.
'@throws ProjectError.ElementNotFound When the column is missing.
Private Sub EnsureControlColumnExists(ByVal colName As String)
    Dim targetName As String

    targetName = IIf(LenB(colName) = 0, COL_CONTROL, colName)
    ColumnIndex targetName
End Sub

'@label:EnsureValidSheetName
'@sub-title Validate that the supplied sheet name is usable
'@param sheetName String. Sheet identifier to validate.
'@throws ProjectError.InvalidArgument When the name is empty or is the header value.
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
'@param infoSelector Byte. Selector from SheetInfoType.
'@return String metadata value.
'@throws ProjectError.InvalidArgument When the selector is unknown.
Private Function SheetInfoValue(ByVal sheetName As String, ByVal infoSelector As Byte) As String
    Dim colIdx As Long
    Dim rowIdx As Long

    rowIdx = EnsureSheetRow(sheetName)

    Select Case infoSelector
        Case LLSheetInfoFieldSheetType
            colIdx = SheetTypeColumnIndex()
        Case LLSheetInfoFieldTableName
            colIdx = TableNameColumnIndex()
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unknown sheet info selector: " & CStr(infoSelector)
    End Select

    SheetInfoValue = CellText(HostSheet().Cells(rowIdx, colIdx).Value)
End Function

'@label:host-sheet
'@prop-title Worksheet holding the dictionary
'@details
'Held on the instance. Reaching it through `dict.Data.Wksh` costs two
'crossings before any work, on every cell this class resolves.
'@return Worksheet
Private Property Get HostSheet() As Worksheet
    If this.hostWksh Is Nothing Then
        EnsureDictionaryReady
        Set this.hostWksh = this.dict.Data.Wksh
    End If
    Set HostSheet = this.hostWksh
End Property

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'A dictionary cell can hold an error value (`#N/A`, `#REF!`) or a Null, and
'`CStr` raises on both.
'@param cellValue Variant. The raw cell value.
'@return String. The text, or an empty string when there is none.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function
    CellText = CStr(cellValue)
End Function

'@label:SheetTypeColumnIndex
'@sub-title Cache the column index for sheet type lookups
'@return Long worksheet column index.
Private Function SheetTypeColumnIndex() As Long
    If this.sheetTypeColumnIndex = 0 Then
        this.sheetTypeColumnIndex = ColumnIndex(COL_SHEET_TYPE)
    End If
    SheetTypeColumnIndex = this.sheetTypeColumnIndex
End Function

'@label:TableNameColumnIndex
'@sub-title Cache the column index for table name lookups and raise if missing
'@return Long worksheet column index.
'@throws ProjectError.ElementNotFound When the table-name column is missing.
Private Function TableNameColumnIndex() As Long
    If this.tableNameColumnIndex = 0 Then
        If Not this.dict.ColumnExists(COL_TABLE_NAME) Then
            ThrowError ProjectError.ElementNotFound, _
                       "Dictionary column '" & COL_TABLE_NAME & "' is required for this operation"
        End If
        this.tableNameColumnIndex = ColumnIndex(COL_TABLE_NAME)
    End If
    TableNameColumnIndex = this.tableNameColumnIndex
End Function

'@label:VariablesProvider
'@sub-title Lazily initialise the LLVariables helper for address resolution
'@return LLVariables cached instance.
Private Function VariablesProvider() As LLVariables
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(this.dict)
    End If
    Set VariablesProvider = this.variables
End Function

'@label:ColumnIndex
'@sub-title Resolve a dictionary column, and raise when it is absent
'@details
'`DataSheet.ColumnIndex` raises with `shouldExist:=True`, so the answer that
'comes back is always a real column.
'@param colName String. Column header to locate.
'@return Long worksheet column index.
'@throws ProjectError.ElementNotFound When the column is absent.
Private Function ColumnIndex(ByVal colName As String) As Long
    ColumnIndex = this.dict.Data.ColumnIndex(colName, shouldExist:=True, matchCase:=False)
End Function

'@section Sheet Row Resolution
'===============================================================================

'@label:sheet-row
'@sub-title Worksheet row that describes one sheet
'@details
'The row is held per sheet name. A held row is checked against the cell it
'points at before it is handed out: `LLdictionary.Prepare` inserts rows, and
'handing out a moved row would answer for the WRONG sheet.
'@param sheetName String. Sheet identifier.
'@return Long. The worksheet row, or 0 when the dictionary has no such sheet.
Private Function ResolveSheetRow(ByVal sheetName As String) As Long
    Dim slot As Long
    Dim rowNum As Long
    Dim found As Range

    If LenB(sheetName) = 0 Then Exit Function

    EnsureDictionaryReady

    slot = CacheSlot(sheetName)
    rowNum = this.cachedRows(slot)

    If rowNum > 0 Then
        If RowStillHolds(rowNum, sheetName) Then
            ResolveSheetRow = rowNum
            Exit Function
        End If

        'The dictionary moved under this object. Forget what this slot holds
        'and search the worksheet again.
        this.cachedRows(slot) = 0
        this.cachedCounts(slot) = -1
        this.cachedFarIndexes(slot) = -1
    End If

    Set found = FindSheetCell(sheetName)
    If found Is Nothing Then Exit Function

    this.cachedRows(slot) = found.Row
    ResolveSheetRow = found.Row
End Function

'@label:row-still-holds
'@sub-title Check one cell against the sheet name it should hold
'@param rowNum Long. Worksheet row to test.
'@param sheetName String. Name the row should carry.
'@return Boolean. True when the cell still reads as sheetName.
Private Function RowStillHolds(ByVal rowNum As Long, ByVal sheetName As String) As Boolean
    Dim offsetRow As Long

    If rowNum <= 0 Then Exit Function

    offsetRow = rowNum - this.sheetRng.Row + 1
    If offsetRow < 1 Then Exit Function
    If offsetRow > this.sheetRng.Rows.Count Then Exit Function

    RowStillHolds = (StrComp(CellText(this.sheetRng.Cells(offsetRow, 1).Value), sheetName, vbTextCompare) = 0)
End Function

'@label:farthest-of
'@sub-title Take the farther of a counted bound and the last written index
'@details
'VarWriter places a variable at the "column index" its own dictionary row
'carries: for an HList sheet that number is the WORKSHEET COLUMN, for a
'VList sheet the WORKSHEET ROW (VarWriter.VarRange). Counting the variables
'and adding the count to the start only agrees with that when the indexes
'run 1, 2, 3 with no gap. A dictionary whose indexes skip a number - the
'shape a setup takes as soon as variables are removed or reordered - writes
'variables past the counted edge, and the data table was then built short of
'them: the columns were on the sheet, outside the ListObject, invisible to
'the filter, the export and every reader that walks the table.
'
'So the counted bound becomes a floor. A dictionary with no index column, or
'one holding nothing numeric, answers with the count as before.
'@param countedBound Long. The bound the variable count gives.
'@param sheetName String. The sheet being measured.
'@return Long. The farther of the two.
Private Function FarthestOf(ByVal countedBound As Long, _
                            ByVal sheetName As String) As Long
    Dim farIndex As Long

    FarthestOf = countedBound
    farIndex = LastVariableIndex(sheetName)
    If farIndex > countedBound Then FarthestOf = farIndex
End Function

'@label:last-variable-index
'@sub-title The largest column index the sheet's variables carry
'@details
'Held per sheet name beside the count, and dropped by the same
'invalidations. A dictionary missing the index column answers 0, which
'leaves every bound where the count put it.
'@param sheetName String. Sheet identifier.
'@return Long. The largest index, or 0 when there is none to read.
Private Function LastVariableIndex(ByVal sheetName As String) As Long
    Dim slot As Long

    slot = CacheSlot(sheetName)
    If this.cachedFarIndexes(slot) < 0 Then
        this.cachedFarIndexes(slot) = MaxColumnIndex(sheetName)
    End If

    LastVariableIndex = this.cachedFarIndexes(slot)
End Function

'@label:max-column-index
'@sub-title Read the index column of one sheet and take its largest number
'@param sheetName String. Sheet identifier.
'@return Long. The largest numeric index, or 0.
Private Function MaxColumnIndex(ByVal sheetName As String) As Long
    Dim indexes As BetterArray
    Dim counter As Long
    Dim cellValue As Variant
    Dim indexValue As Long

    'A dictionary that carries no index column is not a fault here: the
    'count then decides the bound, which is what happened before.
    On Error Resume Next
        Set indexes = this.dict.Data.FilterData(COL_SHEET_NAME, sheetName, _
                                                COL_COLUMN_INDEX, matchCase:=False)
    On Error GoTo 0

    If indexes Is Nothing Then Exit Function

    For counter = indexes.LowerBound To indexes.UpperBound
        cellValue = indexes.Item(counter)
        If IsNumeric(cellValue) Then
            indexValue = CLng(cellValue)
            If indexValue > MaxColumnIndex Then MaxColumnIndex = indexValue
        End If
    Next
End Function

'@label:cache-slot
'@sub-title Find or create the cache slot of one sheet name
'@details
'Names are matched without case, which is the rule the lookups use. A
'dictionary holds a handful of sheets, so a walk over the slots is cheaper
'than anything that would index them.
'@param sheetName String. Sheet identifier.
'@return Long. The slot number, from 1.
Private Function CacheSlot(ByVal sheetName As String) As Long
    Dim counter As Long

    For counter = 1 To this.cachedSlots
        If StrComp(this.cachedNames(counter), sheetName, vbTextCompare) = 0 Then
            CacheSlot = counter
            Exit Function
        End If
    Next

    this.cachedSlots = this.cachedSlots + 1
    ReDim Preserve this.cachedNames(1 To this.cachedSlots)
    ReDim Preserve this.cachedRows(1 To this.cachedSlots)
    ReDim Preserve this.cachedCounts(1 To this.cachedSlots)
    ReDim Preserve this.cachedFarIndexes(1 To this.cachedSlots)

    this.cachedNames(this.cachedSlots) = sheetName
    this.cachedRows(this.cachedSlots) = 0
    this.cachedCounts(this.cachedSlots) = -1
    this.cachedFarIndexes(this.cachedSlots) = -1

    CacheSlot = this.cachedSlots
End Function

'@label:FindSheetCell
'@sub-title Locate the range corresponding to a sheet name
'@details
'`LookIn`, `SearchOrder` and `SearchDirection` are all passed. Excel reuses
'whatever the last `Find` anywhere in the application used when they are left
'out, and that includes the user's own Ctrl+F dialog.
'The header row is refused here, so every caller gets the same rule: the
'sheet-name range is captured with its header, and the header text is a column
'title rather than a sheet.
'@param sheetName String. Sheet identifier (may be empty).
'@return Range when found; Nothing otherwise.
Private Function FindSheetCell(ByVal sheetName As String) As Range
    Dim searchValue As String
    Dim result As Range

    EnsureDictionaryReady

    If LenB(sheetName) = 0 Then Exit Function

    searchValue = SanitizeSearchTerm(sheetName)

    'A Find over a range that holds no match raises on some hosts.
    On Error Resume Next
    Set result = this.sheetRng.Find(What:=searchValue, _
                                    After:=this.sheetRng.Cells(1, 1), _
                                    LookIn:=xlValues, _
                                    LookAt:=xlWhole, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlNext, _
                                    MatchCase:=False)
    On Error GoTo 0

    If Not result Is Nothing Then
        If result.Row = this.sheetRng.Row Then Set result = Nothing
    End If

    Set FindSheetCell = result
End Function

'@label:SanitizeSearchTerm
'@sub-title Escape Excel wildcard characters before using Range.Find
'@details
'Excel reads `*`, `?`, `[`, `]` and `#` in the search text as patterns, so a
'sheet named with one of them would match the wrong row. `LLdictionary` keeps
'its own copy beside the search it serves; a class cannot reach a standard
'module, so the two live where they are used.
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
    this.cachedSlots = 0
    Erase this.cachedNames
    Erase this.cachedRows
    Erase this.cachedCounts
    Erase this.cachedFarIndexes
    Set this.variables = Nothing
    Set this.hostWksh = Nothing
    Set this.sheetRng = Nothing
End Sub

'@label:InvalidateCaches
'@sub-title Drop everything this object holds about the worksheet
'@details
'The way out for a caller that changed the dictionary behind this object. The
'variable counts and the sheet-name range are held for the life of the
'instance, so a caller that adds or removes dictionary rows calls this to be
'answered from the worksheet again.
'@export
Public Sub InvalidateCaches()
    ResetCaches
End Sub

'@label:ValidateDictionary
'@sub-title Ensure the supplied dictionary exposes the required columns
'@param dict LLdictionary to validate.
'@throws ProjectError.ObjectNotInitialized When the dictionary or its DataSheet is missing.
'@throws ProjectError.ElementNotFound When a required column is missing.
Private Sub ValidateDictionary(ByVal dict As LLdictionary)
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
'@param dict LLdictionary to inspect.
'@param colName String. Column header to verify.
'@throws ProjectError.ElementNotFound When the column is missing.
Private Sub EnsureColumnExists(ByVal dict As LLdictionary, ByVal colName As String)
    If Not dict.ColumnExists(colName) Then
        ThrowError ProjectError.ElementNotFound, _
                   "Dictionary column '" & colName & "' is required for LLSheets"
    End If
End Sub

'@section Error Handling
'===============================================================================

'@label:ThrowError
'@sub-title Raise a ProjectError-coded runtime error
'@param errNumb Long. The ProjectError value.
'@param errorMessage String human readable description.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub
