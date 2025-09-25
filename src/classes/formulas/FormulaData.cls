VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Data Used for the Formula Object"

'@PredeclaredId
'@Folder("Formula")
'@ModuleDescription("Provides cached lookups for formula parsing")
'@IgnoreModule UseMeaningfulName, HungarianNotation
'@depends BetterArray

Option Explicit
Implements IFormulaData

'@section State
'===============================================================================

Private Type TFormulaData
    wksh As Worksheet
    functions As BetterArray
    characters As BetterArray
End Type

Private Const CLASSNAME As String = "FormulaData"
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
Private Sub Initialise(ByVal formWksh As Worksheet)
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
End Sub

'@sub-title Build in-memory caches from the worksheet tables
Private Sub BuildCaches()
    Set this.functions = BuildLookup(ResolveColumnRange(FORMULAS_TABLE_NAME, FORMULAS_COLUMN_NAME), True)
    Set this.characters = BuildLookup(ResolveColumnRange(CHARACTERS_TABLE_NAME, CHARACTERS_COLUMN_NAME), False)
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

'@section ErrorHandling
'===============================================================================

'@sub-title Raise a consistent error for this class
'@param errNumb ProjectError describing the failure.
'@param errorMessage String human-readable explanation.
Private Sub ThrowError(ByVal errNumb As ProjectError, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error: " & errorMessage
End Sub

'@section Interface
'===============================================================================

Private Function IFormulaData_SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    IFormulaData_SpecialCharacterIncludes = SpecialCharacterIncludes(specialChar)
End Function

Private Function IFormulaData_ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    IFormulaData_ExcelFormulasIncludes = ExcelFormulasIncludes(ExcelFormula)
End Function

