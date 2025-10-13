VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueOfFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Parse VALUE_OF formulas and resolve lookup metadata"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Parse VALUE_OF formulas and resolve lookup metadata")
'@depends BetterArray, LLdictionary, LLVariables, LLSheets
'@label:ValueOfFormula
'@section Overview
'@details Converts VALUE_OF custom formulas written with dictionary variable names into workbook-ready expressions that supply sheet names and column indices. The class validates inputs against the dictionary and guards against cross-sheet lookups.

Option Explicit
Implements IValueOfFormula

'@section State
'===============================================================================

Private Type TValueOfFormula
    expression As String            'Original VALUE_OF specification
    dict As ILLdictionary           'Dictionary reference providing metadata
    variables As ILLVariables       'Helper for variable lookups
    sheets As ILLSheets             'Helper for sheet validation
    parsed As Boolean               'Tracks whether parsing already occurred
    valid As Boolean                'Final validity flag
    converted As String             'Converted VALUE_OF expression
    failureReason As String         'Human readable failure message
    lookupSheet As String           'Resolved lookup worksheet name
    lookupColumnIndex As Long       'Resolved lookup column index
    valueColumnIndex As Long        'Resolved return column index
    arguments As BetterArray        'Captured argument list for diagnostics
End Type

Private this As TValueOfFormula

Private Const CLASSNAME As String = "ValueOfFormula"
Private Const TOKEN_NAME As String = "VALUE_OF"
Private Const ARGUMENT_SEPARATOR As String = ","
Private Const QUOTE_CHAR As String = """"
Private Const ERROR_ARGUMENT_COUNT As String = "VALUE_OF requires exactly three arguments"
Private Const ERROR_EMPTY_ARGUMENT As String = "VALUE_OF arguments cannot be empty"
Private Const ERROR_VARIABLE_MISSING As String = "Dictionary variable '%1' is not defined"
Private Const ERROR_SHEET_UNRESOLVED As String = "Worksheet entry is missing for lookup variable '%1'"
Private Const ERROR_SHEET_MISMATCH As String = "VALUE_OF expects lookup and value variables to share the same worksheet"
Private Const ERROR_COLUMN_INDEX As String = "Dictionary column index is missing for variable '%1'"
Private Const ERROR_EMPTY_FORMULA As String = "VALUE_OF formula cannot be blank"

'@section Factory
'===============================================================================

'@label:Create
'@sub-title Instantiate a VALUE_OF parser bound to the supplied dictionary
'@param formula String representation of the VALUE_OF call.
'@param dict ILLdictionary exposing variables and sheets metadata.
'@param variables Optional ILLVariables to reuse existing helper instances.
'@param sheets Optional ILLSheets helper to reuse.
'@return IValueOfFormula ready for inspection.
Public Function Create(ByVal formula As String, _
                       ByVal dict As ILLdictionary, _
                       Optional ByVal variables As ILLVariables, _
                       Optional ByVal sheets As ILLSheets) As IValueOfFormula

    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Unable to create ValueOfFormula without a dictionary reference"
    End If

    With New ValueOfFormula
        .Initialise formula, dict, variables, sheets
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================

'@label:Self
'@sub-title Expose the current instance through the interface
'@return IValueOfFormula reference to the current object.
Public Property Get Self() As IValueOfFormula
    Set Self = Me
End Property

'@label:Dictionary
'@sub-title Retrieve the dictionary backing this instance
'@return ILLdictionary used during parsing.
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

'@sub-title Assign the dictionary and reset cached state
'@param dict ILLdictionary providing metadata.
Public Property Set Dictionary(ByVal dict As ILLdictionary)
    If dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "ValueOfFormula requires a valid dictionary reference"
    End If
    Set this.dict = dict
    Set this.variables = Nothing
    Set this.sheets = Nothing
    ResetState
End Property

'@label:ValueOfExpression
'@sub-title Retrieve the stored VALUE_OF expression
'@return String expression as provided to the parser.
Public Property Get ValueOfExpression() As String
    ValueOfExpression = this.expression
End Property

'@sub-title Store the VALUE_OF expression and reset cached results
'@param formula String representing the VALUE_OF call.
Public Property Let ValueOfExpression(ByVal formula As String)
    this.expression = Trim$(formula)
    ResetState
End Property

'@section Validation
'===============================================================================

'@label:Valid
'@sub-title Determine whether the VALUE_OF expression parsed successfully
'@return Boolean True when parsing and dictionary lookups succeeded.
Public Property Get Valid() As Boolean
    EnsureParsed
    Valid = this.valid
End Property

'@label:ConvertedFormula
'@sub-title Retrieve the workbook-ready VALUE_OF expression
'@return String converted expression; vbNullString when invalid.
Public Property Get ConvertedFormula() As String
    EnsureParsed
    ConvertedFormula = this.converted
End Property

'@label:FailureReason
'@sub-title Provide a descriptive message when parsing fails
'@return String description; vbNullString for valid formulas.
Public Property Get FailureReason() As String
    EnsureParsed
    FailureReason = this.failureReason
End Property

'@label:LookupSheetName
'@sub-title Expose the resolved lookup worksheet name
'@return String sheet name or vbNullString when unavailable.
Public Property Get LookupSheetName() As String
    EnsureParsed
    LookupSheetName = this.lookupSheet
End Property

'@label:LookupColumnIndex
'@sub-title Expose the column index used to locate the matching value
'@return Long lookup column index; zero when the formula is invalid.
Public Property Get LookupColumnIndex() As Long
    EnsureParsed
    LookupColumnIndex = this.lookupColumnIndex
End Property

'@label:ValueColumnIndex
'@sub-title Expose the column index providing the returned value
'@return Long value column index; zero when the formula is invalid.
Public Property Get ValueColumnIndex() As Long
    EnsureParsed
    ValueColumnIndex = this.valueColumnIndex
End Property

'@section Parsing
'===============================================================================

'@label:EnsureParsed
'@sub-title Lazily parse the VALUE_OF expression and cache outcomes
Private Sub EnsureParsed()
    Dim args As BetterArray
    Dim lower As Long
    Dim keyVar As String
    Dim lookupVar As String
    Dim valueVar As String
    Dim vars As ILLVariables
    Dim sheetName As String
    Dim sheets As ILLSheets

    If this.parsed Then Exit Sub

    ResetParsedOutcome
    this.parsed = True

    If LenB(ValueOfExpression) = 0 Then
        Fail ERROR_EMPTY_FORMULA
        Exit Sub
    End If

    On Error GoTo UnexpectedFailure

    Set args = ExtractArguments()
    If args Is Nothing Then
        Fail ERROR_ARGUMENT_COUNT
        Exit Sub
    End If

    If args.Length <> 3 Then
        Fail ERROR_ARGUMENT_COUNT
        Exit Sub
    End If

    lower = args.LowerBound
    keyVar = Trim$(CStr(args.Item(lower)))
    lookupVar = Trim$(CStr(args.Item(lower + 1)))
    valueVar = Trim$(CStr(args.Item(lower + 2)))

    If LenB(keyVar) = 0 Or LenB(lookupVar) = 0 Or LenB(valueVar) = 0 Then
        Fail ERROR_EMPTY_ARGUMENT
        Exit Sub
    End If

    Set vars = VariablesProvider()

    If Not vars.Contains(keyVar) Then
        Fail Replace$(ERROR_VARIABLE_MISSING, "%1", keyVar)
        Exit Sub
    End If

    If Not vars.Contains(lookupVar) Then
        Fail Replace$(ERROR_VARIABLE_MISSING, "%1", lookupVar)
        Exit Sub
    End If

    If Not vars.Contains(valueVar) Then
        Fail Replace$(ERROR_VARIABLE_MISSING, "%1", valueVar)
        Exit Sub
    End If

    sheetName = vars.Value("sheet name", lookupVar)
    If LenB(sheetName) = 0 Then
        Fail Replace$(ERROR_SHEET_UNRESOLVED, "%1", lookupVar)
        Exit Sub
    End If

    If StrComp(sheetName, vars.Value("sheet name", valueVar), vbTextCompare) <> 0 Then
        Fail ERROR_SHEET_MISMATCH
        Exit Sub
    End If

    Set sheets = SheetsProvider()
    If Not sheets.Contains(sheetName) Then
        Fail Replace$("Worksheet '%1' referenced by VALUE_OF is missing from the dictionary", "%1", sheetName)
        Exit Sub
    End If

    this.lookupColumnIndex = vars.Index(lookupVar)
    If this.lookupColumnIndex <= 0 Then
        Fail Replace$(ERROR_COLUMN_INDEX, "%1", lookupVar)
        Exit Sub
    End If

    this.valueColumnIndex = vars.Index(valueVar)
    If this.valueColumnIndex <= 0 Then
        Fail Replace$(ERROR_COLUMN_INDEX, "%1", valueVar)
        Exit Sub
    End If

    this.lookupSheet = sheetName
    this.converted = BuildConvertedFormula(keyVar, sheetName, this.lookupColumnIndex, this.valueColumnIndex)
    this.valid = True
    Set this.arguments = args.Clone()
    Exit Sub

UnexpectedFailure:
    Fail "Unable to parse VALUE_OF formula: " & Err.Description
End Sub

'@label:ExtractArguments
'@sub-title Split the VALUE_OF body into trimmed argument segments
'@return BetterArray containing up to three arguments; Nothing when malformed.
Private Function ExtractArguments() As BetterArray
    Dim body As String
    Dim args As BetterArray
    Dim startPos As Long
    Dim idx As Long
    Dim depth As Long
    Dim insideQuotes As Boolean
    Dim currentChar As String
    Dim segment As String

    body = SanitisedFormula()
    If LenB(body) = 0 Then Exit Function

    Set args = New BetterArray
    args.LowerBound = 1
    startPos = 1

    For idx = 1 To Len(body)
        currentChar = Mid$(body, idx, 1)

        Select Case currentChar
            Case QUOTE_CHAR
                insideQuotes = Not insideQuotes
            Case "("
                If Not insideQuotes Then depth = depth + 1
            Case ")"
                If Not insideQuotes And depth > 0 Then depth = depth - 1
            Case ARGUMENT_SEPARATOR
                If Not insideQuotes And depth = 0 Then
                    segment = Trim$(Mid$(body, startPos, idx - startPos))
                    args.Push segment
                    startPos = idx + 1
                End If
        End Select
    Next

    '@Ignore AssignmentNotUsed
    segment = Trim$(Mid$(body, startPos, Len(body) - startPos + 1))
    If LenB(segment) > 0 Then
        args.Push segment
    ElseIf args.Length > 0 Then
        args.Push vbNullString
    End If

    Set ExtractArguments = args
End Function

'@label:SanitisedFormula
'@sub-title Remove the VALUE_OF token and surrounding parenthesis
'@return String containing the argument list; vbNullString when malformed.
Private Function SanitisedFormula() As String
    Dim working As String
    Dim openParen As Long

    working = Trim$(ValueOfExpression)
    If LenB(working) = 0 Then Exit Function

    If StrComp(Left$(UCase$(working), Len(TOKEN_NAME)), TOKEN_NAME, vbBinaryCompare) <> 0 Then Exit Function

    openParen = InStr(working, "(")
    If openParen = 0 Then Exit Function

    working = Mid$(working, openParen + 1)
    working = Trim$(working)
    If LenB(working) = 0 Then Exit Function

    If Right$(working, 1) <> ")" Then Exit Function

    '@Ignore AssignmentNotUsed
    working = Left$(working, Len(working) - 1)
    SanitisedFormula = Trim$(working)
End Function

'@label:BuildConvertedFormula
'@sub-title Compose the converted VALUE_OF expression
'@param keyVar String variable supplying the lookup key range.
'@param sheetName String worksheet name hosting the lookup table.
'@param lookupIndex Long column index for the lookup key.
'@param valueIndex Long column index for the return value.
'@return String workbook-ready VALUE_OF expression.
Private Function BuildConvertedFormula(ByVal keyVar As String, _
                                       ByVal sheetName As String, _
                                       ByVal lookupIndex As Long, _
                                       ByVal valueIndex As Long) As String

    BuildConvertedFormula = TOKEN_NAME & "(" & keyVar & ", " & _
                            QuotedText(sheetName) & ", " & _
                            CStr(lookupIndex) & ", " & _
                            CStr(valueIndex) & ")"
End Function

'@label:QuotedText
'@sub-title Wrap a value in double quotes while escaping embedded quotes
'@param value String to quote.
'@return String suitable for Excel formulas.
Private Function QuotedText(ByVal value As String) As String
    QuotedText = QUOTE_CHAR & Replace$(value, QUOTE_CHAR, QUOTE_CHAR & QUOTE_CHAR) & QUOTE_CHAR
End Function

'@section Helpers
'===============================================================================

'@label:Initialise
'@sub-title Assign dependencies and prime the parser state
Public Sub Initialise(ByVal formula As String, _
                       ByVal dict As ILLdictionary, _
                       ByVal variables As ILLVariables, _
                       ByVal sheets As ILLSheets)
    Set this.dict = dict
    If Not (variables Is Nothing) Then Set this.variables = variables
    If Not (sheets Is Nothing) Then Set this.sheets = sheets
    this.expression = Trim$(formula)
    ResetState
End Sub

'@label:ResetState
'@sub-title Clear cached parsing outcomes
Private Sub ResetState()
    this.parsed = False
    ResetParsedOutcome
    Set this.arguments = Nothing
End Sub

'@label:ResetParsedOutcome
'@sub-title Reset fields populated during parsing
Private Sub ResetParsedOutcome()
    this.valid = False
    this.converted = vbNullString
    this.failureReason = vbNullString
    this.lookupSheet = vbNullString
    this.lookupColumnIndex = 0
    this.valueColumnIndex = 0
End Sub

'@label:Fail
'@sub-title Record a parsing failure
'@param message String describing the failure.
Private Sub Fail(ByVal message As String)
    this.failureReason = message
    this.valid = False
    this.converted = vbNullString
    this.lookupSheet = vbNullString
    this.lookupColumnIndex = 0
    this.valueColumnIndex = 0
End Sub

'@label:VariablesProvider
'@sub-title Lazily create the LLVariables helper
'@return ILLVariables for the current dictionary.
Private Function VariablesProvider() As ILLVariables
    EnsureDictionaryReady
    If this.variables Is Nothing Then
        Set this.variables = LLVariables.Create(this.dict)
    End If
    Set VariablesProvider = this.variables
End Function

'@label:SheetsProvider
'@sub-title Lazily create the LLSheets helper
'@return ILLSheets for the current dictionary.
Private Function SheetsProvider() As ILLSheets
    EnsureDictionaryReady
    If this.sheets Is Nothing Then
        Set this.sheets = LLSheets.Create(this.dict)
    End If
    Set SheetsProvider = this.sheets
End Function

'@label:EnsureDictionaryReady
'@sub-title Guard against missing dictionary references
Private Sub EnsureDictionaryReady()
    If this.dict Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "ValueOfFormula dictionary reference is not initialised"
    End If
End Sub

'@label:ThrowError
'@sub-title Raise a project-specific error
'@param errNumber ProjectError describing the failure.
'@param message String error message.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub

'@section Interface
'===============================================================================

Private Property Get IValueOfFormula_ConvertedFormula() As String
    IValueOfFormula_ConvertedFormula = ConvertedFormula()
End Property

Private Property Get IValueOfFormula_Valid() As Boolean
    IValueOfFormula_Valid = Valid()
End Property

Private Property Get IValueOfFormula_FailureReason() As String
    IValueOfFormula_FailureReason = FailureReason()
End Property

Private Property Get IValueOfFormula_LookupSheetName() As String
    IValueOfFormula_LookupSheetName = LookupSheetName()
End Property

Private Property Get IValueOfFormula_LookupColumnIndex() As Long
    IValueOfFormula_LookupColumnIndex = LookupColumnIndex()
End Property

Private Property Get IValueOfFormula_ValueColumnIndex() As Long
    IValueOfFormula_ValueColumnIndex = ValueColumnIndex()
End Property
