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

'@class ChoiceFormula
'@description
'Converts CHOICE_FORMULA custom functions into CASE_WHEN representations
'and exposes helper metadata for downstream validation. The first argument
'of CHOICE_FORMULA is the choice list name; the remaining arguments form
'condition/result pairs that are delegated to a CaseWhen instance for
'nested IF conversion and category extraction.
'@depends CaseWhen, BetterArray
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Backing UDT, constants, and the answers computed at creation.

Private Type TChoiceFormula
    formula As String               'Original formula specification, trimmed
    body As String                  'Text between the parentheses
    arguments As BetterArray        'The choice name followed by condition/result entries
    resolvedName As String          'First argument: the choice list name
    valid As Boolean                'Whether the specification can be converted
    failureReason As String         'Why the specification was rejected
    caseWhenObject As CaseWhen      'Lazily created CaseWhen wrapper
    sealed As Boolean
End Type

Private this As TChoiceFormula

Private Const CLASSNAME As String = "ChoiceFormula"
Private Const TOKEN_NAME As String = "CHOICE_FORMULA"
Private Const TOKEN_SEPARATOR As String = ","
Private Const QUOTE_CHAR As String = """"

Private Const REASON_EMPTY As String = "The CHOICE_FORMULA expression is empty"
Private Const REASON_TOKEN As String = "The formula does not start with CHOICE_FORMULA("
Private Const REASON_PARENTHESES As String = "CHOICE_FORMULA is missing its opening or closing parenthesis"
Private Const REASON_ARGUMENTS As String = _
    "CHOICE_FORMULA needs a choice list name and at least one condition and one result"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Instantiate a ChoiceFormula parser bound to a formula
'@details
'Trims the incoming formula, stores it in a new instance, and returns
'the ChoiceFormula for fluent downstream operations. The token check, the
'argument split and the choice name are all resolved at assignment, because
'the formula cannot change after Seal. The CASE_WHEN wrapper stays lazy:
'SetupErrors builds one of these per dictionary row and reads the choice
'name alone.
'@param formula String. Representation of a CHOICE_FORMULA call.
'@return ChoiceFormula. A ready-to-query parser instance.
'@export
Public Function Create(ByVal formula As String) As ChoiceFormula
    Dim instance As ChoiceFormula
    Set instance = New ChoiceFormula
    instance.ChoiceSpecifications = formula
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Prevent further changes to setup-only properties
'@details
'Marks the instance as sealed so guarded setters raise when invoked
'after construction. Called by the factory before returning.
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
'@description Properties that expose internal state.

'@label:choice-specifications
'@prop-title Retrieve the stored CHOICE_FORMULA expression
'@details
'Returns the original formula text as stored after trimming during
'assignment. This is the full CHOICE_FORMULA(...) expression. It is the
'read half of the pair the factory writes through.
'@return String. Original formula specification.
Public Property Get ChoiceSpecifications() As String
    ChoiceSpecifications = this.formula
End Property

'@label:choice-specifications-set
'@prop-title Store the CHOICE_FORMULA expression and parse it once
'@details
'Trims the incoming formula and stores it, then runs the token check, the
'argument split and the choice-name resolution right here. Every public
'answer below reads a field afterwards. The old class ran the sanitiser up
'to four times and the validity check once per public call, over a string
'that cannot change after Seal.
'@param formula String. Expression to parse.
Public Property Let ChoiceSpecifications(ByVal formula As String)
    GuardNotSealed "ChoiceSpecifications"
    this.formula = Trim$(formula)
    EvaluateFormula
End Property

'@section Validation
'===============================================================================
'@description Reads the answers computed when the formula was stored.

'@label:valid
'@prop-title Determine whether the stored formula is a CHOICE_FORMULA with arguments
'@details
'Requires the CHOICE_FORMULA token, balanced parentheses, and at least
'three arguments (name plus one condition/result pair). A blank argument is
'dropped by the splitter, so it lowers the argument count: writing
'CHOICE_FORMULA(mylist, , "Yes") leaves two arguments and fails the count.
'@return Boolean. True when the specification can be converted to CASE_WHEN.
'@export
Public Property Get Valid() As Boolean
    Valid = this.valid
End Property

'@label:failure-reason
'@prop-title Explain why the specification was rejected
'@details
'Holds the message for the check that failed, and an empty string when the
'specification is valid. ParsedFormula answers with an empty string for a
'bad expression, and an empty string on its own reads as "no formula was
'written". CaseWhen and ValueOfFormula carry the same property.
'@return String. The failure message, or an empty string.
'@export
Public Property Get FailureReason() As String
    FailureReason = this.failureReason
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
Public Function ParsedFormula() As String
    Dim caseObj As CaseWhen

    If Not this.valid Then Exit Function

    Set caseObj = CaseWhenWrapper()
    If caseObj Is Nothing Then Exit Function

    ParsedFormula = caseObj.ParsedFormula()
End Function

'@label:categories
'@sub-title Extract the result labels from the CHOICE_FORMULA results
'@details
'Delegates to the CaseWhen converter to reuse category parsing logic, and
'hands back what it returns. CaseWhen builds a fresh array on every call, so
'no copy is made here. Returns an empty BetterArray when the formula is
'invalid or the CaseWhen wrapper cannot be created.
'@return BetterArray. Collection of plain category labels.
'@export
Public Function Categories() As BetterArray
    Dim caseObj As CaseWhen

    If this.valid Then
        Set caseObj = CaseWhenWrapper()
        If Not caseObj Is Nothing Then
            Set Categories = caseObj.Categories()
            Exit Function
        End If
    End If

    Set Categories = NewTable()
End Function

'@label:choice-name
'@sub-title Retrieve the name of the choice list referenced by the formula
'@details
'Returns the first argument of the CHOICE_FORMULA expression, which
'identifies the choice list that provides the valid options. It is resolved
'at creation, and it answers even when the specification is invalid: a
'formula that names a list and stops still names a list, and SetupErrors
'reads the name to collect the lists a dictionary uses.
'@return String. Choice list identifier.
'@export
Public Function ChoiceName() As String
    ChoiceName = this.resolvedName
End Function

'@section PrivateHelpers
'===============================================================================
'@description Internal utilities for parsing and CASE_WHEN conversion.

'@label:evaluate-formula
'@sub-title Check the token, split the arguments and resolve the choice name
'@details
'Runs once, from the formula setter. Sets the argument body, the argument
'list, the choice name, the validity answer and the failure message.
Private Sub EvaluateFormula()
    Dim working As String
    Dim openParen As Long

    this.valid = False
    this.failureReason = vbNullString
    this.body = vbNullString
    this.resolvedName = vbNullString
    Set this.arguments = NewTable()
    Set this.caseWhenObject = Nothing

    working = this.formula
    If LenB(working) = 0 Then
        this.failureReason = REASON_EMPTY
        Exit Sub
    End If

    If StrComp(Left$(UCase$(working), Len(TOKEN_NAME)), TOKEN_NAME, vbBinaryCompare) <> 0 Then
        this.failureReason = REASON_TOKEN
        Exit Sub
    End If

    openParen = InStr(working, "(")
    If openParen = 0 Then
        this.failureReason = REASON_PARENTHESES
        Exit Sub
    End If

    'The text before the ( has to be the token in full. Without this second
    'check CHOICE_FORMULAX(...) reads as a CHOICE_FORMULA, because the first
    'check only looks at the leading characters.
    If StrComp(Trim$(Left$(working, openParen - 1)), TOKEN_NAME, vbBinaryCompare) <> 0 Then
        this.failureReason = REASON_TOKEN
        Exit Sub
    End If

    working = Trim$(Mid$(working, openParen + 1))
    If LenB(working) = 0 Then
        this.failureReason = REASON_PARENTHESES
        Exit Sub
    End If

    If Right$(working, 1) <> ")" Then
        this.failureReason = REASON_PARENTHESES
        Exit Sub
    End If

    this.body = Trim$(Left$(working, Len(working) - 1))
    If LenB(this.body) = 0 Then
        this.failureReason = REASON_ARGUMENTS
        Exit Sub
    End If

    Set this.arguments = SplitArguments(this.body)

    'The choice name is resolved even for a specification that goes on to
    'fail the count check, because a formula that names a list still names it.
    If this.arguments.Length > 0 Then
        this.resolvedName = CStr(this.arguments.Item(this.arguments.LowerBound))
    End If

    'Need name + at least one condition/result pair
    If this.arguments.Length < 3 Then
        this.failureReason = REASON_ARGUMENTS
        Exit Sub
    End If

    this.valid = True
End Sub

'@label:new-table
'@sub-title Build an empty one-based BetterArray
'@details
'Every array this class hands out or holds starts at index 1.
'@return BetterArray. An empty array with LowerBound 1.
Private Function NewTable() As BetterArray
    Dim emptyTable As BetterArray

    Set emptyTable = New BetterArray
    emptyTable.LowerBound = 1

    Set NewTable = emptyTable
End Function

'@label:split-arguments
'@sub-title Split the argument body into individual arguments
'@details
'Walks the body character by character, tracking quotation state and
'parenthesis depth. Splits on top-level commas only. Blank segments are
'dropped, so a blank argument lowers the argument count rather than being
'stored.
'CaseWhen.SplitFormulaSegments and ValueOfFormula.ExtractArguments hold the
'same walk; ValueOfFormula keeps its blank segments, because it counts them.
'The three stay separate: a class cannot reach a standard module, so a
'shared copy has nowhere to live.
'@param body String. The comma-delimited text between the parentheses.
'@return BetterArray. The choice name followed by condition/result entries.
Private Function SplitArguments(ByVal body As String) As BetterArray
    Dim arguments As BetterArray
    Dim position As Long
    Dim segmentStart As Long
    Dim depth As Long
    Dim insideQuotes As Boolean
    Dim currentChar As String
    Dim segmentText As String
    Dim formulaLength As Long

    Set arguments = NewTable()

    segmentStart = 1
    formulaLength = Len(body)

    For position = 1 To formulaLength
        currentChar = Mid$(body, position, 1)
        Select Case currentChar
            Case QUOTE_CHAR
                'A doubled quote toggles the flag off and straight back on,
                'so the walker lands in the right state for an escaped quote.
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

    segmentText = Trim$(Mid$(body, segmentStart, formulaLength - segmentStart + 1))
    If LenB(segmentText) > 0 Then arguments.Push segmentText

    Set SplitArguments = arguments
End Function

'@label:case-when-formula
'@sub-title Build the equivalent CASE_WHEN formula from parsed arguments
'@details
'Skips the first argument (choice list name) and joins the remaining
'condition/result entries into a CASE_WHEN(...) expression string.
'Returns an empty string when nothing follows the name.
'The join uses ", ". CaseWhen splits on a top-level comma and trims each
'segment, so the extra space is dropped again. The two have to keep
'agreeing on the separator.
'@return String. CASE_WHEN expression or empty when there is nothing to build.
Private Function CaseWhenFormula() As String
    Dim idx As Long
    Dim lower As Long
    Dim upper As Long
    Dim builder As String

    lower = this.arguments.LowerBound
    upper = this.arguments.UpperBound

    If upper <= lower Then Exit Function          'Missing name or arguments

    builder = "CASE_WHEN("

    For idx = lower + 1 To upper
        builder = builder & CStr(this.arguments.Item(idx))
        If idx < upper Then builder = builder & TOKEN_SEPARATOR & " "
    Next idx

    CaseWhenFormula = builder & ")"
End Function

'@label:case-when-wrapper
'@prop-title Provide a lazily created CaseWhen instance
'@details
'Builds the CASE_WHEN expression from parsed arguments and delegates to
'CaseWhen.Create to obtain the parser. The instance is cached and the same
'one is handed to every caller: CaseWhen is read-only after Seal, so
'sharing it is safe. Making CaseWhen mutable would break this.
'@return CaseWhen. The converted formula parser.
Private Property Get CaseWhenWrapper() As CaseWhen
    Dim caseFormula As String

    If Not (this.caseWhenObject Is Nothing) Then
        Set CaseWhenWrapper = this.caseWhenObject
        Exit Property
    End If

    caseFormula = CaseWhenFormula()
    If LenB(caseFormula) = 0 Then Exit Property

    Set this.caseWhenObject = CaseWhen.Create(caseFormula)
    Set CaseWhenWrapper = this.caseWhenObject
End Property

'@label:throw-error
'@sub-title Raise a project-specific error
'@details
'Wrapper around Err.Raise that standardises the source to the class
'constant CLASSNAME for consistent stack traces. The code is taken as a
'Long: a named enum type in a parameter position is the macOS trap.
'@param errNumber Long. The error code describing the failure.
'@param message String. Error message to propagate.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub
