VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides cached lookups of allowed excel functions for formula parsing"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Provides cached lookups of allowed excel functions for formula parsing")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IFormulaData

'@class FormulaData
'@description
'Provides cached, in-memory lookups for the Excel functions and special
'characters that the formula parser is allowed to accept. On creation it
'reads two ListObjects from a worksheet (T_XlsFonctions for functions,
'T_ascii for separator characters) and stores the values in BetterArray
'caches. It also maintains a registry of grouped formula mappings that
'translate custom tokens (e.g. SUMIFS, MEANIFS) into their canonical
'Excel aggregation functions. Consumers interact with this class
'exclusively through the IFormulaData interface.
'@depends BetterArray
'@version 1.0 (2026-02-09)
'@author OutbreakTools


'@section State
'===============================================================================
'@description Backing UDT and module-level constants for cached lookups.

Private Type TFormulaData
    wksh As Worksheet
    functions As BetterArray
    characters As BetterArray
    groupFunctions As BetterArray        'Registered grouped formula mappings
End Type

Private Const FORMULAS_TABLE_NAME As String = "T_XlsFonctions"
Private Const FORMULAS_COLUMN_NAME As String = "ENG"
Private Const CHARACTERS_TABLE_NAME As String = "T_ascii"
Private Const CHARACTERS_COLUMN_NAME As String = "TEXT"

Private this As TFormulaData

'@section Factory
'===============================================================================

'@label:create
'@sub-title Instantiate a FormulaData helper bound to a worksheet source
'@details
'Validates that the supplied worksheet contains the required ListObjects
'(T_XlsFonctions and T_ascii) with expected columns, then reads their
'data into memory caches. The returned instance exposes its surface
'through the IFormulaData interface so consumers never depend on the
'concrete class.
'@param formWksh Worksheet. The worksheet hosting the lookup tables.
'@return IFormulaData. A fully initialised instance ready for use by the Formulas parser.
'@throws ProjectError.ObjectNotInitialized When formWksh is Nothing.
'@throws ProjectError.ElementNotFound When a required table or column is missing.
'@throws ProjectError.ErrorUnexpectedState When a required column contains no data.
'@depends BetterArray
'@export
Public Function Create(ByVal formWksh As Worksheet) As IFormulaData
    ValidateWorksheet formWksh

    With New FormulaData
        .Initialise formWksh
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================
'@description Properties that expose the internal state and worksheet ranges.

'@label:self
'@prop-title Current object instance
'@details
'Convenience accessor that returns the current instance cast to its
'interface. Used internally by the factory method to hand back an
'IFormulaData reference from the predeclared Create pattern.
'@return IFormulaData. The current instance cast to the interface.
Public Property Get Self() As IFormulaData
    Set Self = Me
End Property

'@label:wksh
'@prop-title Worksheet hosting the lookup tables
'@details
'Returns the worksheet reference stored during creation. All cached
'lookups originate from ListObjects on this worksheet.
'@return Worksheet. The worksheet reference stored at creation.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.wksh
End Property

'@label:wksh-set
'@prop-title Assign the worksheet hosting the lookup tables
'@details
'Replaces the worksheet reference and invalidates all cached lookups
'so they are rebuilt from the new worksheet on the next access.
'@param formWksh Worksheet. The worksheet providing formula metadata.
Public Property Set Wksh(ByVal formWksh As Worksheet)
    Set this.wksh = formWksh
    ResetCaches
End Property

'@label:formula-range
'@prop-title Data body range containing allowed formulas
'@details
'Resolves the ENG column of the T_XlsFonctions ListObject and returns
'its DataBodyRange. Callers can use this range to inspect the raw
'worksheet data backing the function whitelist.
'@return Range. The data body range covering the ENG column of the formulas table.
'@throws ProjectError.ElementNotFound When the table or column cannot be resolved.
Public Property Get FormulaRange() As Range
    Set FormulaRange = ResolveColumnRange(FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME)
End Property

'@label:special-characters-range
'@prop-title Data body range containing allowed separator characters
'@details
'Resolves the TEXT column of the T_ascii ListObject and returns its
'DataBodyRange. Provides direct access to the characters that the
'tokeniser treats as delimiters.
'@return Range. The data body range covering the TEXT column of the characters table.
'@throws ProjectError.ElementNotFound When the table or column cannot be resolved.
Public Property Get SpecialCharactersRange() As Range
    Set SpecialCharactersRange = ResolveColumnRange(CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME)
End Property

'@section Lookups
'===============================================================================
'@description Membership tests against the cached function and character lists.

'@label:special-character-includes
'@sub-title Determine whether a special character is allowed
'@details
'Checks the in-memory character cache for the provided token. Returns
'False immediately when the input is empty. The cache is built lazily
'from the T_ascii table on first access.
'@param specialChar String. Single-character token to inspect.
'@return Boolean. True when the character is listed in the lookup table.
'@export
Private Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    Dim characterLookup As BetterArray

    If LenB(specialChar) = 0 Then Exit Function

    Set characterLookup = EnsureCharacters()
    SpecialCharacterIncludes = characterLookup.Includes(specialChar)
End Function

'@label:excel-formulas-includes
'@sub-title Determine whether an Excel function is allowed
'@details
'Normalises the input to uppercase and checks the in-memory function
'cache. Returns False immediately when the input is empty. The cache
'is built lazily from the T_XlsFonctions table on first access.
'@param ExcelFormula String. Function name to inspect.
'@return Boolean. True when the function is listed in the lookup table.
'@export
Private Function ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    Dim functionLookup As BetterArray
    Dim normalized As String

    If LenB(ExcelFormula) = 0 Then Exit Function

    Set functionLookup = EnsureFunctions
    normalized = UCase$(ExcelFormula)
    ExcelFormulasIncludes = functionLookup.Includes(normalized)
End Function

'@section PrivateHelpers
'===============================================================================
'@description Initialisation, validation, and cache management utilities.

'@label:initialise
'@sub-title Perform one-time initialisation of the instance
'@details
'Stores the worksheet reference, clears any stale caches, and eagerly
'builds both function and character lookups so the instance is ready
'for immediate use after creation.
'@param formWksh Worksheet. The worksheet providing lookup sources.
Public Sub Initialise(ByVal formWksh As Worksheet)
    Set this.wksh = formWksh
    ResetCaches
    BuildCaches
End Sub

'@label:validate-worksheet
'@sub-title Ensure the worksheet contains the required list objects and columns
'@details
'Guards against Nothing references and delegates to ValidateListObject
'for each of the two required tables (formulas and characters). Raises
'an error at the first missing element encountered.
'@param targetSheet Worksheet. The worksheet to validate.
'@throws ProjectError.ObjectNotInitialized When targetSheet is Nothing.
'@throws ProjectError.ElementNotFound When a required table or column is missing.
'@throws ProjectError.ErrorUnexpectedState When a column contains no data.
Private Sub ValidateWorksheet(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "No worksheet supplied for FormulaData"
    End If

    ValidateListObject targetSheet, FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME
    ValidateListObject targetSheet, CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME
End Sub

'@label:validate-list-object
'@sub-title Confirm the presence of a specific ListObject and column
'@details
'Attempts to locate the named ListObject and column on the host
'worksheet using guarded On Error Resume Next blocks. Raises a
'ProjectError when any component is missing or the column has no data.
'@param hostSheet Worksheet. The worksheet containing the table.
'@param tableName String. Name of the ListObject to inspect.
'@param columnName String. Header expected within the table.
'@throws ProjectError.ElementNotFound When the table or column is absent.
'@throws ProjectError.ErrorUnexpectedState When the column DataBodyRange is Nothing.
Private Sub ValidateListObject(ByVal hostSheet As Worksheet, _
                               ByVal tableName As String, _
                               ByVal columnName As String)
    Dim tableObj As ListObject
    Dim targetColumn As ListColumn

    On Error Resume Next
        Set tableObj = hostSheet.ListObjects(tableName)
    On Error GoTo 0

    If tableObj Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Table " & tableName & " was not found on worksheet " & hostSheet.Name
    End If

    On Error Resume Next
        Set targetColumn = tableObj.ListColumns(columnName)
    On Error GoTo 0

    If targetColumn Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Column " & columnName & " was not found in table " & tableName
    End If

    If targetColumn.DataBodyRange Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                   "Table " & tableName & " column " & columnName & " does not contain any data"
    End If
End Sub

'@label:resolve-column-range
'@sub-title Resolve and return the data body range for a column
'@details
'Looks up the named table and column on the stored worksheet and
'returns its DataBodyRange. Raises an error when the range cannot be
'resolved, which typically indicates the table structure was modified
'after the instance was created.
'@param tableName String. ListObject name.
'@param columnName String. ListColumn header.
'@return Range. The data body range containing the column data.
'@throws ProjectError.ElementNotFound When the column cannot be resolved.
Private Function ResolveColumnRange(ByVal tableName As String, _
                                    ByVal columnName As String) As Range
    Dim columnRange As Range

    On Error Resume Next
        Set columnRange = Wksh.ListObjects(tableName).ListColumns(columnName).DataBodyRange
    On Error GoTo 0

    If columnRange Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Unable to resolve column " & columnName & " in table " & tableName
    End If

    Set ResolveColumnRange = columnRange
End Function

'@label:reset-caches
'@sub-title Reset cached lookups so they reload on next use
'@details
'Sets the three cache references (functions, characters, group
'functions) to Nothing. The next call to any lookup method will
'trigger a lazy rebuild from the worksheet.
Private Sub ResetCaches()
    Set this.functions = Nothing
    Set this.characters = Nothing
    Set this.groupFunctions = Nothing
End Sub

'@label:build-caches
'@sub-title Build in-memory caches from the worksheet tables
'@details
'Reads the formula and character ranges into BetterArray instances
'and seeds the grouped-function registry with default mappings.
'Function names are uppercased for case-insensitive matching while
'character tokens are stored verbatim.
Private Sub BuildCaches()
    Set this.functions = BuildLookup(ResolveColumnRange(FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME), True)
    Set this.characters = BuildLookup(ResolveColumnRange(CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME), False)
    InitialiseGroupFunctions
End Sub

'@label:ensure-functions
'@sub-title Ensure the function lookup is initialised and return it
'@details
'Lazily rebuilds all caches when the function BetterArray is Nothing,
'then returns the function lookup for immediate use.
'@return BetterArray. Collection keyed by uppercase function names.
Private Function EnsureFunctions() As BetterArray
    If this.functions Is Nothing Then BuildCaches
    Set EnsureFunctions = this.functions
End Function

'@label:ensure-characters
'@sub-title Ensure the character lookup is initialised and return it
'@details
'Lazily rebuilds all caches when the character BetterArray is Nothing,
'then returns the character lookup for immediate use.
'@return BetterArray. Collection keyed by literal characters.
Private Function EnsureCharacters() As BetterArray
    If this.characters Is Nothing Then BuildCaches
    Set EnsureCharacters = this.characters
End Function

'@label:build-lookup
'@sub-title Convert a column of cells into a lookup array
'@details
'Iterates over each cell in the source range, skips empty values,
'optionally uppercases the entry, and pushes only distinct values
'into the resulting BetterArray. The array uses a zero-based lower
'bound for consistent indexing across all caches.
'@param sourceRange Range. The cells containing the values to cache.
'@param normaliseCase Boolean. True to uppercase entries before storing.
'@return BetterArray. Collection containing the distinct values.
Private Function BuildLookup(ByVal sourceRange As Range, _
                             ByVal normaliseCase As Boolean) As BetterArray
    Dim list As BetterArray
    Dim cell As Range
    Dim entry As String

    Set list = New BetterArray
    list.LowerBound = 0

    For Each cell In sourceRange.Cells
        entry = CStr(cell.Value)
        If LenB(entry) > 0 Then
            If normaliseCase Then entry = UCase$(entry)
            If Not list.Includes(entry) Then list.Push entry
        End If
    Next cell

    Set BuildLookup = list
End Function

'@section GroupFunctions
'===============================================================================
'@description Maintains an in-memory registry of grouped formulas along with
'the canonical Excel aggregation function each token maps to.

'@label:initialise-group-functions
'@sub-title Seed the grouped-formula registry with recognised mappings
'@details
'Ensures the registry exists and registers the seven built-in grouped
'formulas (SUMIFS, COUNTIFS, NIFS, MINIFS, MAXIFS, MEDIANIFS, MEANIFS)
'together with their corresponding Excel aggregators and native flags.
Private Sub InitialiseGroupFunctions()

    EnsureGroupRegistry

    RegisterGroupFunction "SUMIFS", "SUMIFS", True
    RegisterGroupFunction "COUNTIFS", "COUNTIFS", True
    RegisterGroupFunction "NIFS", "COUNTIFS", True
    RegisterGroupFunction "MINIFS", "MIN", False
    RegisterGroupFunction "MAXIFS", "MAX", False
    RegisterGroupFunction "MEDIANIFS", "MEDIAN", False
    RegisterGroupFunction "MEANIFS", "AVERAGE", False
End Sub

'@label:ensure-group-registry
'@sub-title Guarantee the grouped-formula registry exists and is seeded
'@details
'Creates a new zero-based BetterArray when the registry is Nothing.
'Called before any read or write operation on the grouped-function list
'to prevent Nothing-reference errors.
Private Sub EnsureGroupRegistry()
    If this.groupFunctions Is Nothing Then
        Set this.groupFunctions = New BetterArray
        this.groupFunctions.LowerBound = 0
    End If
End Sub

'@label:register-group-function
'@sub-title Register a grouped formula token and its Excel aggregator
'@details
'Normalises both names to uppercase, removes any prior mapping for the
'same token, and stores a three-element array (token, aggregator,
'useNative flag). Also ensures that both the token and its aggregator
'appear in the main function cache so the tokeniser accepts them.
'@param functionName String. Grouped token as it appears in setup formulas.
'@param aggregatorName String. Canonical Excel aggregator to generate.
'@param useNative Optional Boolean. When True the output uses the native *IFS function directly. Defaults to False.
Private Sub RegisterGroupFunction(ByVal functionName As String, _
                                  ByVal aggregatorName As String, _
                                  Optional ByVal useNative As Boolean = False)
    Dim normalizedFunction As String
    Dim normalizedAggregator As String
    Dim mapping As Variant

    normalizedFunction = UCase$(Trim$(functionName))
    normalizedAggregator = UCase$(Trim$(aggregatorName))

    If LenB(normalizedFunction) = 0 Then Exit Sub
    If LenB(normalizedAggregator) = 0 Then Exit Sub

    EnsureGroupRegistry
    RemoveGroupFunctionMapping normalizedFunction

    mapping = Array(normalizedFunction, normalizedAggregator, CBool(useNative))
    this.groupFunctions.Push mapping

    EnsureFunctionToken normalizedFunction
    EnsureFunctionToken normalizedAggregator
End Sub

'@label:ensure-function-token
'@sub-title Insert a token into the Excel-function cache when absent
'@details
'Normalises the token to uppercase, creates the function cache if
'needed, and appends the token only when it is not already present.
'This ensures grouped function names are accepted by the tokeniser.
'@param functionToken String. Token to guarantee in the lookup.
Private Sub EnsureFunctionToken(ByVal functionToken As String)
    Dim normalized As String

    normalized = UCase$(Trim$(functionToken))
    If LenB(normalized) = 0 Then Exit Sub

    If this.functions Is Nothing Then
        Set this.functions = New BetterArray
        this.functions.LowerBound = 0
    End If

    If Not this.functions.Includes(normalized) Then
        this.functions.Push normalized
    End If
End Sub

'@label:remove-group-function-mapping
'@sub-title Remove any existing mapping for the provided grouped token
'@details
'Iterates the registry in reverse order and removes entries whose
'first element matches the normalised function name. Reverse iteration
'prevents index shifting issues when removing from a BetterArray.
'@param functionName String. Grouped formula token to clear.
Private Sub RemoveGroupFunctionMapping(ByVal functionName As String)
    Dim idx As Long
    Dim entry As Variant
    Dim normalized As String

    If this.groupFunctions Is Nothing Then Exit Sub
    If this.groupFunctions.Length = 0 Then Exit Sub

    normalized = UCase$(Trim$(functionName))
    If LenB(normalized) = 0 Then Exit Sub

    For idx = this.groupFunctions.UpperBound To this.groupFunctions.LowerBound Step -1
        entry = this.groupFunctions.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalized, vbBinaryCompare) = 0 Then
                this.groupFunctions.Remove idx
            End If
        End If
    Next idx
End Sub

'@label:group-function-includes
'@sub-title Determine whether a grouped function token is registered
'@details
'Normalises the input and scans the registry for a matching first
'element. Returns True on the first match. Returns False when the
'registry is empty or no match is found.
'@param functionName String. Grouped token to look up.
'@return Boolean. True when the token is known.
'@export
Private Function GroupFunctionIncludes(ByVal functionName As String) As Boolean
    Dim idx As Long
    Dim entry As Variant
    Dim normalized As String

    normalized = UCase$(Trim$(functionName))
    If LenB(normalized) = 0 Then Exit Function

    EnsureGroupRegistry

    If this.groupFunctions.Length = 0 Then Exit Function

    For idx = this.groupFunctions.LowerBound To this.groupFunctions.UpperBound
        entry = this.groupFunctions.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalized, vbBinaryCompare) = 0 Then
                GroupFunctionIncludes = True
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:group-aggregator
'@sub-title Retrieve the Excel aggregator linked to a grouped formula token
'@details
'Scans the registry for the matching token and returns the second
'element of the mapping array, which is the canonical Excel aggregator
'function name. Returns vbNullString when the token is not found.
'@param functionName String. Grouped token to translate.
'@return String. Aggregator name or vbNullString when unknown.
'@export
Private Function GroupAggregator(ByVal functionName As String) As String
    Dim idx As Long
    Dim entry As Variant
    Dim normalized As String

    normalized = UCase$(Trim$(functionName))
    If LenB(normalized) = 0 Then Exit Function

    EnsureGroupRegistry

    If this.groupFunctions.Length = 0 Then Exit Function

    For idx = this.groupFunctions.LowerBound To this.groupFunctions.UpperBound
        entry = this.groupFunctions.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalized, vbBinaryCompare) = 0 Then
                GroupAggregator = CStr(entry(1))
                Exit Function
            End If
        End If
    Next idx
End Function

'@label:group-uses-native
'@sub-title Determine whether the grouped function should emit a native Excel *IFS aggregator
'@details
'Looks up the third element of the mapping array for the given token.
'When True, the Formulas class emits the native *IFS function (e.g.
'SUMIFS) instead of wrapping the aggregator in an IF block.
'@param functionName String. Grouped token to inspect.
'@return Boolean. True when the output uses the corresponding *IFS function directly.
'@export
Private Function GroupUsesNative(ByVal functionName As String) As Boolean
    Dim idx As Long
    Dim entry As Variant
    Dim normalized As String

    normalized = UCase$(Trim$(functionName))
    If LenB(normalized) = 0 Then Exit Function

    EnsureGroupRegistry

    If this.groupFunctions.Length = 0 Then Exit Function

    For idx = this.groupFunctions.LowerBound To this.groupFunctions.UpperBound
        entry = this.groupFunctions.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), normalized, vbBinaryCompare) = 0 Then
                GroupUsesNative = CBool(entry(2))
                Exit Function
            End If
        End If
    Next idx
End Function

'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising helper for consistent ProjectError usage.

'@label:throw-error
'@sub-title Raise a consistent error for this class
'@details
'Wrapper around Err.Raise that standardises the source to "FormulaData",
'providing a consistent stack trace across all methods in this class.
'@param errNumb Long. ProjectError code describing the failure.
'@param errorMessage String. Human-readable explanation.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "FormulaData", "Error: " & errorMessage
End Sub

'@section Interface
'===============================================================================
'@description Delegated members satisfying the IFormulaData contract.
'See the corresponding Private members above for full documentation.

Private Function IFormulaData_SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    IFormulaData_SpecialCharacterIncludes = SpecialCharacterIncludes(specialChar)
End Function

Private Function IFormulaData_ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    IFormulaData_ExcelFormulasIncludes = ExcelFormulasIncludes(ExcelFormula)
End Function

Private Function IFormulaData_IsGroupFunction(ByVal functionName As String) As Boolean
    IFormulaData_IsGroupFunction = GroupFunctionIncludes(functionName)
End Function

Private Function IFormulaData_GroupAggregator(ByVal functionName As String) As String
    IFormulaData_GroupAggregator = GroupAggregator(functionName)
End Function

Private Function IFormulaData_GroupUsesNativeFunction(ByVal functionName As String) As Boolean
    IFormulaData_GroupUsesNativeFunction = GroupUsesNative(functionName)
End Function