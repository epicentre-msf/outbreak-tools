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

'@class FormulaData
'@description
'Answers what the formula parser is allowed to accept. Create reads two
'ListObjects from a worksheet, T_XlsFonctions for the Excel function names and
'T_ascii for the separator characters, and holds each of them as one delimited
'string. ExcelFormulasIncludes says whether a token is a known Excel function,
'and SpecialCharacterIncludes whether a character is a known separator.
'
'GROUPED FORMULAS
'-------------------------------------------------------------------------------
'Seven tokens stand for an aggregation over rows that match a set of criteria,
'SUMIFS and MEANIFS among them. IsGroupFunction says whether a token is one of
'the seven, GroupAggregator names the Excel function it stands for, and
'GroupUsesNativeFunction says whether the output uses the matching Excel *IFS
'function directly.
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Backing UDT and module-level constants for cached lookups.

Private Type TFormulaData
    wksh As Worksheet
    functionIndex As String         'Uppercase function names, delimited
    characterIndex As String        'Separator characters as written, delimited
    sealed As Boolean               'Set by the factory, closes the setter
End Type

Private Const CLASSNAME As String = "FormulaData"
Private Const FORMULAS_TABLE_NAME As String = "T_XlsFonctions"
Private Const FORMULAS_COLUMN_NAME As String = "ENG"
Private Const CHARACTERS_TABLE_NAME As String = "T_ascii"
Private Const CHARACTERS_COLUMN_NAME As String = "TEXT"

'Both lookups are held as one string with this character around every entry,
'so a membership test is one InStr rather than a walk over a BetterArray.
'Formulas asks SpecialCharacterIncludes once per character of every formula in
'the workbook and ExcelFormulasIncludes once per token, and BetterArray.Includes
'copies the whole internal array before it searches. A worksheet cell cannot
'hold this character, so it can never be part of an entry.
Private Const INDEX_SEPARATOR As String = vbNullChar

Private this As TFormulaData

'@section Factory
'===============================================================================

'@label:create
'@sub-title Instantiate a FormulaData helper bound to a worksheet source
'@details
'Validates that the supplied worksheet contains the required ListObjects
'(T_XlsFonctions and T_ascii) with expected columns, then reads their
'data into memory. Each table is resolved once: the old code validated the
'two tables and then resolved both again to read them.
'@param formWksh Worksheet. The worksheet hosting the lookup tables.
'@return FormulaData. A fully initialised instance ready for use by the Formulas parser.
'@throws ProjectError.ObjectNotInitialized When formWksh is Nothing.
'@throws ProjectError.ElementNotFound When a required table or column is missing.
'@throws ProjectError.ErrorUnexpectedState When a required column contains no data.
'@export
Public Function Create(ByVal formWksh As Worksheet) As FormulaData
    If formWksh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "No worksheet supplied for FormulaData"
    End If

    Dim instance As FormulaData
    Set instance = New FormulaData
    instance.Initialise formWksh
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Prevent further changes to setup-only properties
'@details
'Marks the instance as sealed so guarded setters raise when invoked after
'construction. Called by the factory before returning. It also closes the
'back door the class used to carry: a direct New followed by Initialise
'skipped nothing, but a later Wksh write rebuilt every lookup underneath a
'parser that was already reading them.
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Reject writes to setup-only properties after sealing
'@details
'Raises an error when a guarded setter is invoked on a sealed instance.
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@section PublicAccessors
'===============================================================================
'@description Properties that expose the internal state.

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
'Replaces the worksheet reference and rereads both lookups from it. The
'factory seals the instance, so this raises after construction.
'@param formWksh Worksheet. The worksheet providing formula metadata.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Public Property Set Wksh(ByVal formWksh As Worksheet)
    GuardNotSealed "Wksh"
    Initialise formWksh
End Property

'@section Lookups
'===============================================================================
'@description Membership tests against the cached function and character lists.

'@label:special-character-includes
'@sub-title Determine whether a special character is allowed
'@details
'Checks the character index for the provided token. Returns False when the
'input is empty. The match is exact, including case: the T_ascii entries are
'stored as they are written on the worksheet.
'@param specialChar String. Single-character token to inspect.
'@return Boolean. True when the character is listed in the lookup table.
'@export
Public Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    If LenB(specialChar) = 0 Then Exit Function

    SpecialCharacterIncludes = IndexHolds(this.characterIndex, specialChar)
End Function

'@label:excel-formulas-includes
'@sub-title Determine whether an Excel function is allowed
'@details
'Normalises the input to uppercase and checks the function index. Returns
'False when the input is empty. Function names are stored uppercase, so the
'lookup is case-insensitive.
'@param ExcelFormula String. Function name to inspect.
'@return Boolean. True when the function is listed in the lookup table.
'@export
Public Function ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    If LenB(ExcelFormula) = 0 Then Exit Function

    ExcelFormulasIncludes = IndexHolds(this.functionIndex, UCase$(ExcelFormula))
End Function

'@section GroupFunctions
'===============================================================================
'@description The grouped formulas and the canonical Excel aggregation function
'each token maps to.

'@label:group-function-includes
'@sub-title Determine whether a grouped function token is registered
'@details
'Normalises the input and looks it up in the mapping. Returns False for an
'empty string and for a token that is not one of the seven.
'@param functionName String. Grouped token to look up.
'@return Boolean. True when the token is known.
'@export
Public Function IsGroupFunction(ByVal functionName As String) As Boolean
    Dim aggregator As String
    Dim usesNative As Boolean

    IsGroupFunction = GroupMappingFound(functionName, aggregator, usesNative)
End Function

'@label:group-aggregator
'@sub-title Retrieve the Excel aggregator linked to a grouped formula token
'@details
'Returns the canonical Excel aggregator function name for the token.
'Returns vbNullString when the token is not one of the seven.
'@param functionName String. Grouped token to translate.
'@return String. Aggregator name or vbNullString when unknown.
'@export
Public Function GroupAggregator(ByVal functionName As String) As String
    Dim aggregator As String
    Dim usesNative As Boolean

    If GroupMappingFound(functionName, aggregator, usesNative) Then
        GroupAggregator = aggregator
    End If
End Function

'@label:group-uses-native
'@sub-title Determine whether the grouped function should emit a native Excel *IFS aggregator
'@details
'When True, the Formulas class emits the native *IFS function (e.g. SUMIFS)
'instead of wrapping the aggregator in an IF block. Returns False when the
'token is not one of the seven.
'@param functionName String. Grouped token to inspect.
'@return Boolean. True when the output uses the corresponding *IFS function directly.
'@export
Public Function GroupUsesNativeFunction(ByVal functionName As String) As Boolean
    Dim aggregator As String
    Dim usesNative As Boolean

    If GroupMappingFound(functionName, aggregator, usesNative) Then
        GroupUsesNativeFunction = usesNative
    End If
End Function

'@label:group-mapping-found
'@sub-title Resolve one grouped token to its aggregator and native flag
'@details
'The seven mappings are compiled in. The old class held them in a
'BetterArray registry with routines to add and remove entries, and nothing
'ever added one: the seven literals were the whole set. The three public
'answers above each walked that registry with the same twenty lines.
'@param functionName String. Grouped token to look up.
'@param aggregator ByRef String. Filled with the Excel aggregator.
'@param usesNative ByRef Boolean. Filled with the native-function flag.
'@return Boolean. True when the token is one of the seven.
Private Function GroupMappingFound(ByVal functionName As String, _
                                   ByRef aggregator As String, _
                                   ByRef usesNative As Boolean) As Boolean
    Dim normalized As String

    aggregator = vbNullString
    usesNative = False

    normalized = UCase$(Trim$(functionName))
    If LenB(normalized) = 0 Then Exit Function

    Select Case normalized
        Case "SUMIFS"
            aggregator = "SUMIFS"
            usesNative = True
        Case "COUNTIFS"
            aggregator = "COUNTIFS"
            usesNative = True
        Case "NIFS"
            aggregator = "COUNTIFS"
            usesNative = True
        Case "MINIFS"
            aggregator = "MIN"
        Case "MAXIFS"
            aggregator = "MAX"
        Case "MEDIANIFS"
            aggregator = "MEDIAN"
        Case "MEANIFS"
            aggregator = "AVERAGE"
        Case Else
            Exit Function
    End Select

    GroupMappingFound = True
End Function

'@label:register-group-tokens
'@sub-title Put the grouped tokens and their aggregators in the function index
'@details
'The seven tokens and the four aggregators they name are accepted by the
'tokeniser whether or not the worksheet table lists them.
'Formulas.AppendToken depends on this and a test asserts it.
Private Sub RegisterGroupTokens()
    AppendFunctionToken "SUMIFS"
    AppendFunctionToken "COUNTIFS"
    AppendFunctionToken "NIFS"
    AppendFunctionToken "MINIFS"
    AppendFunctionToken "MAXIFS"
    AppendFunctionToken "MEDIANIFS"
    AppendFunctionToken "MEANIFS"
    AppendFunctionToken "MIN"
    AppendFunctionToken "MAX"
    AppendFunctionToken "MEDIAN"
    AppendFunctionToken "AVERAGE"
End Sub

'@section PrivateHelpers
'===============================================================================
'@description Initialisation, validation, and index building.

'@label:initialise
'@sub-title Perform one-time initialisation of the instance
'@details
'Stores the worksheet reference and reads both lookup columns into memory.
'It stays Public because the factory calls it on a second instance, and it
'is guarded: after the factory seals that instance, this raises.
'@param formWksh Worksheet. The worksheet providing lookup sources.
'@throws ProjectError.ElementNotFound When a required table or column is missing.
'@throws ProjectError.ErrorUnexpectedState When a required column contains no data.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Public Sub Initialise(ByVal formWksh As Worksheet)
    GuardNotSealed "Initialise"

    Set this.wksh = formWksh
    this.functionIndex = BuildIndex(ColumnRange(FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME), True)
    this.characterIndex = BuildIndex(ColumnRange(CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME), False)
    RegisterGroupTokens
End Sub

'@label:column-range
'@sub-title Resolve the data body range of one lookup column
'@details
'Locates the named ListObject and column on the stored worksheet and
'returns the DataBodyRange, raising the error that says which part is
'missing. This is the only place either table is resolved.
'@param tableName String. ListObject name.
'@param columnName String. ListColumn header.
'@return Range. The data body range containing the column data.
'@throws ProjectError.ElementNotFound When the table or the column is absent.
'@throws ProjectError.ErrorUnexpectedState When the column holds no data.
Private Function ColumnRange(ByVal tableName As String, _
                             ByVal columnName As String) As Range
    Dim tableObj As ListObject
    Dim targetColumn As ListColumn

    'Scoped: Excel raises when a ListObject or a ListColumn is absent, and
    'the message it gives names neither the table nor the sheet.
    On Error Resume Next
        Set tableObj = this.wksh.ListObjects(tableName)
    On Error GoTo 0

    If tableObj Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Table " & tableName & " was not found on worksheet " & this.wksh.Name
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

    Set ColumnRange = targetColumn.DataBodyRange
End Function

'@label:build-index
'@sub-title Read one column and turn it into a delimited lookup string
'@details
'Reads the whole column in one crossing, then walks the values in memory.
'The old code read one cell at a time and called BetterArray.Includes once
'per row against the very list it was filling, which copies the whole
'internal array on every call. Duplicates are skipped with one InStr over
'the string being built.
'@param sourceRange Range. The cells containing the values to index.
'@param normaliseCase Boolean. True to uppercase entries before storing.
'@return String. The distinct values, each wrapped in INDEX_SEPARATOR.
Private Function BuildIndex(ByVal sourceRange As Range, _
                            ByVal normaliseCase As Boolean) As String
    Dim entries As Variant
    Dim entry As String
    Dim counter As Long
    Dim builder As String

    entries = ColumnValues(sourceRange)
    builder = INDEX_SEPARATOR

    For counter = LBound(entries) To UBound(entries)
        entry = entries(counter)
        If LenB(entry) > 0 Then
            If normaliseCase Then entry = UCase$(entry)
            If InStr(1, builder, INDEX_SEPARATOR & entry & INDEX_SEPARATOR, vbBinaryCompare) = 0 Then
                builder = builder & entry & INDEX_SEPARATOR
            End If
        End If
    Next

    BuildIndex = builder
End Function

'@label:column-values
'@sub-title Read a one-column range into a string array
'@details
'A range of one cell answers with a scalar rather than an array, so that
'case is handled on its own.
'@param sourceRange Range. The one-column range to read.
'@return Variant. A one-based array of strings.
Private Function ColumnValues(ByVal sourceRange As Range) As Variant
    Dim values As Variant
    Dim result() As String
    Dim counter As Long

    If sourceRange.Cells.Count = 1 Then
        ReDim result(1 To 1)
        result(1) = CellText(sourceRange.Value)
    Else
        values = sourceRange.Value
        ReDim result(1 To UBound(values, 1))
        For counter = 1 To UBound(values, 1)
            result(counter) = CellText(values(counter, 1))
        Next
    End If

    ColumnValues = result
End Function

'@label:cell-text
'@sub-title Read one cell value as text
'@details
'A cell holding an error value raises a type mismatch on CStr, and a table
'the user is editing can hold one.
'@param cellValue Variant. The value read from the worksheet.
'@return String. The text, or an empty string.
Private Function CellText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function

    CellText = CStr(cellValue)
End Function

'@label:index-holds
'@sub-title Test one entry against a delimited lookup string
'@param lookupIndex String. The delimited index to search.
'@param entry String. The entry to look for.
'@return Boolean. True when the index holds the entry.
Private Function IndexHolds(ByVal lookupIndex As String, ByVal entry As String) As Boolean
    IndexHolds = (InStr(1, lookupIndex, INDEX_SEPARATOR & entry & INDEX_SEPARATOR, vbBinaryCompare) > 0)
End Function

'@label:append-function-token
'@sub-title Add one token to the function index when it is not already there
'@param functionToken String. Token to guarantee in the lookup.
Private Sub AppendFunctionToken(ByVal functionToken As String)
    Dim normalized As String

    normalized = UCase$(Trim$(functionToken))
    If LenB(normalized) = 0 Then Exit Sub

    If Not IndexHolds(this.functionIndex, normalized) Then
        this.functionIndex = this.functionIndex & normalized & INDEX_SEPARATOR
    End If
End Sub

'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising helper for consistent ProjectError usage.

'@label:throw-error
'@sub-title Raise a consistent error for this class
'@details
'Wrapper around Err.Raise that standardises the source to "FormulaData",
'providing a consistent stack trace across all methods in this class. The
'message goes through as written, the same as the other five classes in
'this folder.
'@param errNumb Long. ProjectError code describing the failure.
'@param errorMessage String. Human-readable explanation.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub
