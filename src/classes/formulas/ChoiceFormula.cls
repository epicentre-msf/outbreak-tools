VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChoiceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Parse and expose CHOICE_FORMULA expressions"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Parse and expose CHOICE_FORMULA expressions")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IChoiceFormula

'@class ChoiceFormula
'@description
'Converts CHOICE_FORMULA custom functions into CASE_WHEN representations
'and exposes helper metadata for downstream validation. The first argument
'of CHOICE_FORMULA is the choice list name; the remaining arguments form
'condition/result pairs that are delegated to a CaseWhen instance for
'nested IF conversion and category extraction.
'@depends CaseWhen, BetterArray
'@version 1.0 (2026-02-09)
'@author Yves Amevoin


'@section State
'===============================================================================
'@description Backing UDT, constants, and lazily populated caches.

Private Type TChoiceFormula
    formula As String              'Original formula specification
    cachedArguments As BetterArray  'Cached list of arguments extracted from the formula
    cachedCaseWhen As ICaseWhen     'Lazily created CaseWhen wrapper
    cachedChoiceName As String      'Memoised choice list identifier
End Type

Private this As TChoiceFormula

Private Const CLASSNAME As String = "ChoiceFormula"
Private Const TOKEN_NAME As String = "CHOICE_FORMULA"
Private Const TOKEN_SEPARATOR As String = ","
Private Const QUOTE_CHAR As String = """"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Instantiate a ChoiceFormula parser bound to a formula
'@details
'Trims the incoming formula, stores it in a new instance, and returns
'the IChoiceFormula interface for fluent downstream operations. No
'eager parsing occurs; the formula is parsed lazily on first access.
'@param formula String. Representation of a CHOICE_FORMULA call.
'@return IChoiceFormula. A ready-to-query parser instance.
'@export
Public Function Create(ByVal formula As String) As IChoiceFormula
    With New ChoiceFormula
        .ChoiceSpecifications = formula
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================
'@description Properties that expose internal state and the interface pointer.

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to its interface. Used internally
'by the factory method to return an IChoiceFormula reference.
'@return IChoiceFormula. Reference to the current object.
Public Property Get Self() As IChoiceFormula
    Set Self = Me
End Property

'@label:choice-specifications
'@prop-title Retrieve the stored CHOICE_FORMULA expression
'@details
'Returns the original formula text as stored after trimming during
'assignment. This is the full CHOICE_FORMULA(...) expression.
'@return String. Original formula specification.
Public Property Get ChoiceSpecifications() As String
    ChoiceSpecifications = this.formula
End Property

'@label:choice-specifications-set
'@prop-title Store the CHOICE_FORMULA expression and reset caches
'@details
'Trims the incoming formula and stores it, then invalidates all
'cached artefacts so subsequent reads trigger a fresh parse.
'@param formula String. Expression to parse.
Public Property Let ChoiceSpecifications(ByVal formula As String)
    this.formula = Trim$(formula)
    ResetCaches
End Property

'@section Validation
'===============================================================================
'@description Checks whether the stored formula matches the CHOICE_FORMULA pattern.

'@label:valid
'@prop-title Determine whether the stored formula is a CHOICE_FORMULA with arguments
'@details
'Requires the CHOICE_FORMULA token, balanced parentheses, and at least
'three arguments (name plus one condition/result pair). Returns False
'when any argument is blank.
'@return Boolean. True when the specification can be converted to CASE_WHEN.
'@export
Private Property Get Valid() As Boolean
    Dim arguments As BetterArray
    Dim argumentCount As Long
    Dim sanitized As String

    sanitized = SanitisedFormula()
    If LenB(sanitized) = 0 Then Exit Property

    Set arguments = ChoiceArguments()
    If arguments Is Nothing Then Exit Property

    argumentCount = arguments.Length
    If argumentCount < 3 Then Exit Property      'Need name + at least one condition/result pair

    If ContainsEmptyArgument(arguments) Then Exit Property

    Valid = True
End Property

'@section Conversion
'===============================================================================
'@description Core conversion, category extraction, and choice name resolution.

'@label:parsed-formula
'@sub-title Convert CHOICE_FORMULA into nested IF statements through CASE_WHEN
'@details
'Builds the equivalent CASE_WHEN expression from the parsed arguments,
'then delegates to the CaseWhen class to produce the nested IF output.
'Returns an empty string when the formula is invalid or conversion fails.
'@return String. Expression ready for worksheet consumption.
'@export
Private Function parsedFormula() As String
    Dim caseObj As ICaseWhen

    If Not Valid() Then Exit Function

    Set caseObj = CaseWhenWrapper()
    If caseObj Is Nothing Then Exit Function

    parsedFormula = caseObj.parsedFormula()
End Function

'@label:categories
'@sub-title Extract textual labels from the CHOICE_FORMULA results
'@details
'Delegates to the CaseWhen converter to reuse category parsing logic.
'Returns an empty BetterArray when the formula is invalid or the
'CaseWhen wrapper cannot be created.
'@return BetterArray. Collection containing quoted category labels.
'@export
Private Function Categories() As BetterArray
    Dim caseObj As ICaseWhen
    Dim categoriesTable As BetterArray

    Set categoriesTable = New BetterArray
    categoriesTable.LowerBound = 1

    If Not Valid() Then
        Set Categories = categoriesTable.Clone()
        Exit Function
    End If

    Set caseObj = CaseWhenWrapper()
    If caseObj Is Nothing Then
        Set Categories = categoriesTable.Clone()
        Exit Function
    End If

    '@Ignore AssignmentNotUsed
    Set categoriesTable = caseObj.Categories()
    Set Categories = categoriesTable.Clone()
End Function

'@label:choice-name
'@sub-title Retrieve the name of the choice list referenced by the formula
'@details
'Returns the first argument of the CHOICE_FORMULA expression, which
'identifies the choice list that provides the valid options. The value
'is memoised after the first resolution.
'@return String. Choice list identifier.
'@export
Private Function choiceName() As String
    choiceName = ResolvedChoiceName()
End Function

'@section PrivateHelpers
'===============================================================================
'@description Internal utilities for cache management, argument parsing, and CASE_WHEN conversion.

'@label:reset-caches
'@sub-title Reset cached artefacts when the formula changes
'@details
'Clears the argument cache, the CaseWhen wrapper, and the memoised
'choice name so the next access triggers a full reparse.
Private Sub ResetCaches()
    Set this.cachedArguments = Nothing
    Set this.cachedCaseWhen = Nothing
    this.cachedChoiceName = vbNullString
End Sub

'@label:sanitised-formula
'@sub-title Resolve the portion of the formula between parentheses
'@details
'Strips the CHOICE_FORMULA token, the opening parenthesis, and the
'closing parenthesis to expose the raw comma-delimited argument body.
'Returns an empty string when the formula does not match the expected
'pattern.
'@return String. The comma-delimited arguments.
Private Function SanitisedFormula() As String
    Dim working As String
    Dim openParen As Long

    working = Trim$(ChoiceSpecifications)
    If LenB(working) = 0 Then Exit Function

    If StrComp(Left$(UCase$(working), Len(TOKEN_NAME)), TOKEN_NAME, vbBinaryCompare) <> 0 Then Exit Function

    openParen = InStr(working, "(")
    If openParen = 0 Then Exit Function

    If StrComp(Trim$(Left$(working, openParen - 1)), TOKEN_NAME, vbBinaryCompare) <> 0 Then Exit Function

    working = Mid$(working, openParen + 1)
    working = Trim$(working)
    If LenB(working) = 0 Then Exit Function

    If Right$(working, 1) <> ")" Then Exit Function

    '@Ignore AssignmentNotUsed
    working = Left$(working, Len(working) - 1)
    SanitisedFormula = Trim$(working)
End Function

'@label:choice-arguments
'@sub-title Split the sanitised formula into individual arguments
'@details
'Walks the sanitised body character by character, tracking quotation
'state and parenthesis depth. Splits on top-level commas only. The
'result is cached so subsequent calls avoid reparsing.
'@return BetterArray. The choice name followed by condition/result entries.
Private Function ChoiceArguments() As BetterArray
    Dim arguments As BetterArray
    Dim body As String
    Dim position As Long
    Dim segmentStart As Long
    Dim depth As Long
    Dim insideQuotes As Boolean
    Dim currentChar As String
    Dim segmentText As String
    Dim formulaLength As Long

    If Not (this.cachedArguments Is Nothing) Then
        Set ChoiceArguments = this.cachedArguments.Clone()
        Exit Function
    End If

    body = SanitisedFormula()
    Set arguments = New BetterArray
    arguments.LowerBound = 1

    If LenB(body) = 0 Then
        Set this.cachedArguments = arguments.Clone()
        Set ChoiceArguments = arguments.Clone()
        Exit Function
    End If

    segmentStart = 1
    formulaLength = Len(body)

    For position = 1 To formulaLength
        currentChar = Mid$(body, position, 1)
        Select Case currentChar
            Case QUOTE_CHAR
                insideQuotes = Not insideQuotes
            Case "("
                If Not insideQuotes Then depth = depth + 1
            Case ")"
                If Not insideQuotes And depth > 0 Then depth = depth - 1
            Case TOKEN_SEPARATOR
                If Not insideQuotes And depth = 0 Then
                    segmentText = Trim$(Mid$(body, segmentStart, position - segmentStart))
                    If LenB(segmentText) > 0 Then arguments.Push segmentText
                    segmentStart = position + 1
                End If
        End Select
    Next position

    '@Ignore AssignmentNotUsed
    segmentText = Trim$(Mid$(body, segmentStart, formulaLength - segmentStart + 1))
    If LenB(segmentText) > 0 Then arguments.Push segmentText

    Set this.cachedArguments = arguments.Clone()
    Set ChoiceArguments = this.cachedArguments.Clone()
End Function

'@label:contains-empty-argument
'@sub-title Determine whether the arguments contain empty entries
'@details
'Scans the BetterArray for any blank entry. Returns True on the first
'empty string encountered.
'@param arguments BetterArray. List of parsed arguments.
'@return Boolean. True when an argument is blank.
Private Function ContainsEmptyArgument(ByVal arguments As BetterArray) As Boolean
    Dim idx As Long
    Dim lower As Long
    Dim upper As Long

    If arguments Is Nothing Then Exit Function

    lower = arguments.LowerBound
    upper = arguments.UpperBound
    If upper < lower Then Exit Function

    For idx = lower To upper
        If LenB(CStr(arguments.Item(idx))) = 0 Then
            ContainsEmptyArgument = True
            Exit Function
        End If
    Next idx
End Function

'@label:case-when-formula
'@sub-title Build the equivalent CASE_WHEN formula from parsed arguments
'@details
'Skips the first argument (choice list name) and joins the remaining
'condition/result entries into a CASE_WHEN(...) expression string.
'Returns an empty string when fewer than two arguments exist beyond
'the name.
'@return String. CASE_WHEN expression or empty when invalid.
Private Function CaseWhenFormula() As String
    Dim arguments As BetterArray
    Dim idx As Long
    Dim lower As Long
    Dim upper As Long
    Dim builder As String

    Set arguments = ChoiceArguments()
    If arguments Is Nothing Then Exit Function

    lower = arguments.LowerBound
    upper = arguments.UpperBound

    If upper <= lower Then Exit Function          'Missing name or arguments
    If (upper - lower) < 1 Then Exit Function     'No arguments beyond the choice name

    builder = "CASE_WHEN("

    For idx = lower + 1 To upper
        builder = builder & CStr(arguments.Item(idx))
        If idx < upper Then builder = builder & TOKEN_SEPARATOR & " "
    Next idx

    '@Ignore AssignmentNotUsed
    builder = builder & ")"
    CaseWhenFormula = builder
End Function

'@label:case-when-wrapper
'@prop-title Provide a lazily created CaseWhen instance
'@details
'Builds the CASE_WHEN expression from parsed arguments and delegates
'to CaseWhen.Create to obtain the parser. The result is cached for
'subsequent accesses. Raises an error when conversion fails.
'@return ICaseWhen. The converted formula parser.
'@throws ProjectError.ErrorUnexpectedState When conversion to CASE_WHEN fails.
Private Property Get CaseWhenWrapper() As ICaseWhen
    Dim caseFormula As String

    If Not (this.cachedCaseWhen Is Nothing) Then
        Set CaseWhenWrapper = this.cachedCaseWhen
        Exit Property
    End If

    caseFormula = CaseWhenFormula()
    If LenB(caseFormula) = 0 Then Exit Property

    On Error GoTo ConversionError
        Set this.cachedCaseWhen = CaseWhen.Create(caseFormula)
    On Error GoTo 0

    Set CaseWhenWrapper = this.cachedCaseWhen
    Exit Property

ConversionError:
    ThrowError ProjectError.ErrorUnexpectedState, _
               "Unable to convert CHOICE_FORMULA to CASE_WHEN: " & Err.Description
End Property

'@label:resolved-choice-name
'@sub-title Resolve and cache the choice list name
'@details
'Extracts the first argument from the parsed argument list, which
'is the choice list identifier. The value is memoised so subsequent
'calls avoid re-parsing.
'@return String. Identifier extracted from the first argument.
Private Function ResolvedChoiceName() As String
    Dim arguments As BetterArray
    Dim lower As Long

    If LenB(this.cachedChoiceName) > 0 Then
        ResolvedChoiceName = this.cachedChoiceName
        Exit Function
    End If

    Set arguments = ChoiceArguments()
    If arguments Is Nothing Then Exit Function

    lower = arguments.LowerBound
    If arguments.UpperBound < lower Then Exit Function

    this.cachedChoiceName = CStr(arguments.Item(lower))
    ResolvedChoiceName = this.cachedChoiceName
End Function

'@label:throw-error
'@sub-title Raise a project-specific error
'@details
'Wrapper around Err.Raise that standardises the source to the class
'constant CLASSNAME for consistent stack traces.
'@param errNumber ProjectError. The error code describing the failure.
'@param message String. Error message to propagate.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub

'@section Interface
'===============================================================================
'@description Delegated members satisfying the IChoiceFormula contract.

Private Property Get IChoiceFormula_parsedFormula() As String
    IChoiceFormula_parsedFormula = parsedFormula()
End Property

Private Property Get IChoiceFormula_Categories() As BetterArray
    Set IChoiceFormula_Categories = Categories()
End Property

Private Property Get IChoiceFormula_Valid() As Boolean
    IChoiceFormula_Valid = Valid()
End Property

Private Property Get IChoiceFormula_ChoiceName() As String
    IChoiceFormula_ChoiceName = choiceName()
End Property
