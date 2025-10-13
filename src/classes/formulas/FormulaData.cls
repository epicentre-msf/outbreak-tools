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
'@depends BetterArray
Option Explicit
Implements IFormulaData

'@section State
'===============================================================================

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

'@label:Create
'@jump:Create
'@sub-title Instantiate a FormulaData helper bound to a worksheet source
'@details Reads the backing ListObjects once and caches allowed tokens in memory.
'@param formWksh Worksheet hosting the lookup tables.
'@return IFormulaData instance ready for use by Formulas.
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

'@label:Self
'@jump:Self
'@sub-title Return the current helper as its interface
'@return IFormulaData reference to this instance.
Public Property Get Self() As IFormulaData
    Set Self = Me
End Property

'@label:Worksheet
'@jump:Worksheet
'@sub-title Retrieve the worksheet hosting the lookup tables
'@return Worksheet reference stored at creation.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.wksh
End Property

'@sub-title Store the worksheet hosting the lookup tables
'@param formWksh Worksheet providing formula metadata.
Public Property Set Wksh(ByVal formWksh As Worksheet)
    Set this.wksh = formWksh
    ResetCaches
End Property

'@label:FormulaRange
'@jump:FormulaRange
'@sub-title Retrieve the data body range containing allowed formulas
'@return Range covering the ENG column of the formulas table.
Public Property Get FormulaRange() As Range
    Set FormulaRange = ResolveColumnRange(FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME)
End Property

'@label:SpecialCharactersRange
'@jump:SpecialCharactersRange
'@sub-title Retrieve the data body range containing allowed characters
'@return Range covering the TEXT column of the characters table.
Public Property Get SpecialCharactersRange() As Range
    Set SpecialCharactersRange = ResolveColumnRange(CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME)
End Property

'@section Lookups
'===============================================================================

'@sub-title Determine whether a special character is allowed
'@param specialChar String single-character token to inspect.
'@return Boolean True when the character is listed in the lookup table.
Private Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    Dim characterLookup As BetterArray

    If LenB(specialChar) = 0 Then Exit Function

    Set characterLookup = EnsureCharacters()
    SpecialCharacterIncludes = characterLookup.Includes(specialChar)
End Function

'@sub-title Determine whether an Excel function is allowed
'@param ExcelFormula String function name to inspect.
'@return Boolean True when the function is listed in the lookup table.
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

'@sub-title Perform one-time initialisation of the instance
'@param formWksh Worksheet providing lookup sources.
Public Sub Initialise(ByVal formWksh As Worksheet)
    Set this.wksh = formWksh
    ResetCaches
    BuildCaches
End Sub

'@sub-title Ensure the worksheet contains the required list objects and columns
'@param targetSheet Worksheet to validate.
Private Sub ValidateWorksheet(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "No worksheet supplied for FormulaData"
    End If

    ValidateListObject targetSheet, FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME
    ValidateListObject targetSheet, CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME
End Sub

'@sub-title Confirm the presence of a specific ListObject and column
'@param hostSheet Worksheet containing the table.
'@param tableName String name of the ListObject to inspect.
'@param columnName String header expected within the table.
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

'@sub-title Resolve and return the data body range for a column
'@param tableName ListObject name.
'@param columnName ListColumn header.
'@return Range containing the column data.
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

'@sub-title Reset cached lookups so they reload on next use
Private Sub ResetCaches()
    Set this.functions = Nothing
    Set this.characters = Nothing
    Set this.groupFunctions = Nothing
End Sub

'@sub-title Build in-memory caches from the worksheet tables
Private Sub BuildCaches()
    Set this.functions = BuildLookup(ResolveColumnRange(FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME), True)
    Set this.characters = BuildLookup(ResolveColumnRange(CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME), False)
    InitialiseGroupFunctions
End Sub

'@sub-title Ensure the function lookup is initialised and return it
'@return BetterArray keyed by uppercase function names.
Private Function EnsureFunctions() As BetterArray
    If this.functions Is Nothing Then BuildCaches
    Set EnsureFunctions = this.functions
End Function

'@sub-title Ensure the character lookup is initialised and return it
'@return BetterArray keyed by literal characters.
Private Function EnsureCharacters() As BetterArray
    If this.characters Is Nothing Then BuildCaches
    Set EnsureCharacters = this.characters
End Function

'@sub-title Convert a column of cells into a lookup array
'@param sourceRange Range containing the values to cache.
'@param normaliseCase Boolean True to uppercase entries before storing.
'@return BetterArray containing the distinct values.
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
'@details Maintains an in-memory registry of grouped formulas along with the
'@details canonical Excel aggregation function each token maps to.

'@label:InitialiseGroupFunctions
'@sub-title Seed the grouped-formula registry with recognised mappings
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

'@label:EnsureGroupRegistry
'@sub-title Guarantee the grouped-formula registry exists and is seeded
Private Sub EnsureGroupRegistry()
    If this.groupFunctions Is Nothing Then
        Set this.groupFunctions = New BetterArray
        this.groupFunctions.LowerBound = 0
    End If
End Sub

'@label:RegisterGroupFunction
'@sub-title Register a grouped formula token and its Excel aggregator
'@param functionName String grouped token as it appears in setup formulas.
'@param aggregatorName String canonical Excel aggregator to generate.
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

'@label:EnsureFunctionToken
'@sub-title Insert a token into the Excel-function cache when absent
'@param functionToken String token to guarantee in the lookup.
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

'@label:RemoveGroupFunctionMapping
'@sub-title Remove any existing mapping for the provided grouped token
'@param functionName String grouped formula token to clear.
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

'@label:GroupFunctionIncludes
'@sub-title Determine whether a grouped function token is registered
'@param functionName String grouped token to look up.
'@return Boolean True when the token is known.
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

'@label:GroupAggregator
'@sub-title Retrieve the Excel aggregator linked to a grouped formula token
'@param functionName String grouped token to translate.
'@return String aggregator name or vbNullString when unknown.
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

'@label:GroupUsesNative
'@sub-title Determine whether the grouped function should emit a native Excel *IFS aggregator
'@param functionName String grouped token to inspect.
'@return Boolean True when the output uses the corresponding *IFS function directly.
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

'@sub-title Raise a consistent error for this class
'@param errNumb ProjectError describing the failure.
'@param errorMessage String human-readable explanation.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "FormulaData", "Error: " & errorMessage
End Sub

'@section Interface
'===============================================================================

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
