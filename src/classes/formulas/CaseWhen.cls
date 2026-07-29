VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseWhen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Handle CaseWhen custom function"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Handle CaseWhen custom function")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class CaseWhen
'@description
'Parses CASE_WHEN custom formulas written in setup expressions and rewrites
'them as nested Excel IF statements. The class splits the formula body into
'alternating condition/result pairs while respecting quoted segments and
'nested parentheses, then builds the equivalent nested IF expression. It
'also extracts the plain result labels for downstream analysis consumers.
'@depends BetterArray
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Backing UDT, constants, and cached parsing table.

Private Type TCaseWhen
    formula As String
    body As String
    parsingTable As BetterArray
    valid As Boolean
    failureReason As String
    sealed As Boolean
End Type

Private this As TCaseWhen

Private Const CLASS_NAME As String = "CaseWhen"
Private Const TOKEN_HEADER As String = "CASE_WHEN("
Private Const TOKEN_SEPARATOR As String = ","
Private Const QUOTE_CHAR As String = """"

Private Const REASON_EMPTY As String = "The CASE_WHEN formula is empty"
Private Const REASON_TOKEN As String = "The formula does not start with CASE_WHEN("
Private Const REASON_NO_BODY As String = "CASE_WHEN() holds no condition and no result"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Instantiate a CaseWhen parser bound to a formula
'@details
'Trims the incoming formula, stores it in a new instance, and returns
'the new instance for fluent downstream consumption.
'The header check runs once, at assignment. The argument table is parsed
'lazily on first access to ParsedFormula or Categories.
'@param formula String. The CASE_WHEN expression to parse.
'@return CaseWhen. A ready-to-query parser instance.
'@export
Public Function Create(ByVal formula As String) As CaseWhen
    Dim instance As CaseWhen
    Set instance = New CaseWhen
    instance.CaseWhenFormula = formula
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

'@label:throw-error
'@sub-title Raise a project-specific error
'@details
'Wrapper around Err.Raise that standardises the source to "CaseWhen".
'@param errNumb Long. ProjectError code describing the failure.
'@param errorMessage String. Human-readable explanation.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASS_NAME, errorMessage
End Sub

'@section PublicAccessors
'===============================================================================
'@description Properties that expose the internal state.

'@label:formula
'@prop-title Retrieve the raw CASE_WHEN formula
'@details
'Returns the original formula text as stored after trimming during
'assignment. This is the full CASE_WHEN(...) expression.
'@return String. Original formula text.
Public Property Get CaseWhenFormula() As String
    CaseWhenFormula = this.formula
End Property

'@label:formula-set
'@prop-title Store the CASE_WHEN formula and check it once
'@details
'Trims the incoming formula, stores it, and runs the header check and the
'body extraction right here. Both answers are held, so Valid, ParsedFormula
'and Categories each read a field. The formula cannot change after Seal, so
'the answers cannot go stale.
'@param formula String. Expression to analyse.
Public Property Let CaseWhenFormula(ByVal formula As String)
    GuardNotSealed "CaseWhenFormula"
    this.formula = Trim$(formula)
    EvaluateFormula
End Property

'@section Validation
'===============================================================================
'@description Reads the answers computed when the formula was stored.

'@label:valid
'@prop-title Determine whether the stored formula is a CASE_WHEN
'@details
'True when the formula begins with the CASE_WHEN( token (case-insensitive)
'and carries content beyond the header. A bare "CASE_WHEN(" is rejected:
'it produces an empty body, and an empty parsed formula reaches the caller
'as "the formula is empty" for an expression the user did write.
'@return Boolean. True when the formula is parseable.
'@export
Public Property Get Valid() As Boolean
    Valid = this.valid
End Property

'@label:failure-reason
'@prop-title Explain why the formula was rejected
'@details
'Holds the message for the check that failed, and an empty string when the
'formula is valid. ParsedFormula answers with an empty string for a bad
'expression, and an empty string on its own reads as "no formula was
'written". The caller reads this to tell the user what really happened.
'ValueOfFormula carries the same property.
'@return String. The failure message, or an empty string.
'@export
Public Property Get FailureReason() As String
    FailureReason = this.failureReason
End Property

'@section ParsingInfrastructure
'===============================================================================
'@description Lazily computed and cached argument table.

'@label:case-when-table
'@prop-title Retrieve the cached list of CASE_WHEN arguments
'@details
'Splits the formula body into alternating condition/result entries while
'respecting quoted segments and nested parentheses. The result is cached so
'later reads avoid reparsing the string. The cached array itself is handed
'back: both readers, ParsedFormula and Categories, only ask for Length and
'Item and neither writes to it.
'@return BetterArray. The ordered CASE_WHEN arguments.
Private Property Get CaseWhenTable() As BetterArray
    If this.parsingTable Is Nothing Then
        If this.valid Then
            Set this.parsingTable = SplitFormulaSegments(this.body)
        Else
            Set this.parsingTable = NewTable()
        End If
    End If

    Set CaseWhenTable = this.parsingTable
End Property

'@section Parsing
'===============================================================================
'@description Core conversion and category extraction logic.

'@label:parsed-formula
'@sub-title Convert CASE_WHEN into nested IF statements
'@details
'Iterates through the argument table in pairs (condition, result) and
'builds Excel-compatible nested IF fragments. When the argument count
'is odd the final entry becomes the default branch; when even an empty
'string literal is appended as the fallback. Excel supports up to 64
'nested IF calls and nothing here counts them.
'@return String. Nested IF expression ready for worksheet consumption.
'@export
Public Function ParsedFormula() As String

    Dim counter As Long
    Dim formula As String
    Dim parenthesis As String
    Dim parsingTable As BetterArray
    Dim totalItems As Long

    If Not this.valid Then Exit Function

    Set parsingTable = CaseWhenTable()
    totalItems = parsingTable.Length
    If totalItems = 0 Then Exit Function

    counter = 1

    'Convert to nested IFs.
    Do While counter < totalItems
        'Emit "IF(condition, result, " fragments and track closing parenthesis
        formula = formula & "IF(" & parsingTable.Item(counter) & ", " & _
                             parsingTable.Item(counter + 1) & ", "
        counter = counter + 2
        parenthesis = parenthesis & ")"
    Loop

    If counter = totalItems Then
        'Odd number of entries => final default branch is present
        formula = formula & parsingTable.Item(counter) & parenthesis
    Else
        'Even number => no default provided, return empty string
        formula = formula & QUOTE_CHAR & QUOTE_CHAR & parenthesis
    End If

    ParsedFormula = formula
End Function

'@label:categories
'@sub-title Extract the result labels from the CASE_WHEN arguments
'@details
'Reads each result expression (even-indexed entries in the argument table)
'and returns the label as plain text, with no quote characters around it.
'That is the shape both consumers want: SetupErrors compares the labels
'against the choice sheet, which holds plain text, and EventSetup draws
'them as dropdown and chart labels.
'@return BetterArray. Collection of plain category labels.
'@export
Public Function Categories() As BetterArray
    Dim categoriesTable As BetterArray
    Dim counter As Long
    Dim parsingTable As BetterArray
    Dim levelValue As String
    Dim totalItems As Long

    Set categoriesTable = NewTable()

    If this.valid Then
        counter = 2
        Set parsingTable = CaseWhenTable()
        totalItems = parsingTable.Length

        Do While counter <= totalItems
            'Result expressions sit at even indexes (2,4,6...)
            levelValue = PlainLabel(CStr(parsingTable.Item(counter)))
            If LenB(levelValue) > 0 Then categoriesTable.Push levelValue

            'The step drops to 1 on the entry before the last one, which is how
            'the trailing default branch of an odd-length table is picked up.
            'Both the even and the odd case depend on this.
            If (counter = (totalItems - 1)) Then
                counter = counter + 1
            Else
                counter = counter + 2
            End If
        Loop
    End If

    Set Categories = categoriesTable
End Function

'@section PrivateHelpers
'===============================================================================
'@description Internal utilities for validation, labels and segmentation.

'@label:evaluate-formula
'@sub-title Check the header and pull out the argument body
'@details
'Runs once, from the formula setter. Sets the validity answer, the failure
'message and the argument body, and drops any parsed table. The old code
'asked this question four times per use — from Valid, from CaseWhenTable,
'from ParsedFormula and from Categories — over a string that cannot change
'after Seal.
Private Sub EvaluateFormula()
    this.valid = False
    this.failureReason = vbNullString
    this.body = vbNullString
    Set this.parsingTable = Nothing

    If LenB(this.formula) = 0 Then
        this.failureReason = REASON_EMPTY
        Exit Sub
    End If

    'The header match ignores case: the setup sheets carry both spellings.
    If StrComp(Left$(this.formula, Len(TOKEN_HEADER)), TOKEN_HEADER, vbTextCompare) <> 0 Then
        this.failureReason = REASON_TOKEN
        Exit Sub
    End If

    this.body = SanitisedFormula()
    If LenB(this.body) = 0 Then
        this.failureReason = REASON_NO_BODY
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

'@label:plain-label
'@sub-title Return one result expression as plain text
'@details
'A quoted result loses its outer pair of quotes, and a doubled inner quote
'becomes one quote — a result written "He said ""hi""" comes back as
'He said "hi". A result with no quotes around it is returned as written.
'@param levelText String. One entry of the argument table.
'@return String. The label as the user will read it.
Private Function PlainLabel(ByVal levelText As String) As String
    Dim innerText As String

    If Len(levelText) >= 2 Then
        If Left$(levelText, 1) = QUOTE_CHAR And Right$(levelText, 1) = QUOTE_CHAR Then
            innerText = Mid$(levelText, 2, Len(levelText) - 2)
            PlainLabel = Replace$(innerText, QUOTE_CHAR & QUOTE_CHAR, QUOTE_CHAR)
            Exit Function
        End If
    End If

    PlainLabel = levelText
End Function

'@label:sanitised-formula
'@sub-title Remove the CASE_WHEN header and closing parenthesis
'@details
'Strips the leading CASE_WHEN( token and the trailing closing parenthesis
'to expose the raw comma-delimited argument body. One trailing bracket is
'removed without checking the balance, so CASE_WHEN(a, b)) keeps one of
'them. The segment walker below survives that, so it is left alone.
'@return String. Inner content of the CASE_WHEN call.
Private Function SanitisedFormula() As String
    Dim trimmedFormula As String

    'The formula is trimmed on the way in, so it is read straight here.
    trimmedFormula = Mid$(this.formula, Len(TOKEN_HEADER) + 1)

    If Right$(trimmedFormula, 1) = ")" Then
        trimmedFormula = Left$(trimmedFormula, Len(trimmedFormula) - 1)
    End If

    SanitisedFormula = Trim$(trimmedFormula)
End Function

'@label:split-formula-segments
'@sub-title Split the CASE_WHEN body into trimmed argument segments
'@details
'Walks the formula body character by character, tracking quotation state
'and parenthesis depth. Splits on top-level commas only, preserving nested
'expressions and quoted strings intact. Blank segments are dropped.
'ChoiceFormula.ChoiceArguments and ValueOfFormula.ExtractArguments hold the
'same walk; ValueOfFormula keeps its blank segments, because it counts them.
'The three stay separate: a class cannot reach a standard module, so a
'shared copy has nowhere to live.
'The walk reads one character at a time. A CASE_WHEN body is a short string
'and this runs once per instance, so the reading cost sits under the cost of
'the string joins around it.
'@param formulaBody String. The text between CASE_WHEN( and the final ).
'@return BetterArray. Alternating condition and result expressions.
Private Function SplitFormulaSegments(ByVal formulaBody As String) As BetterArray
    Dim segments As BetterArray
    Dim segmentStart As Long
    Dim position As Long
    Dim depth As Long
    Dim insideQuotes As Boolean
    Dim currentChar As String
    Dim segmentText As String

    Set segments = NewTable()

    segmentStart = 1

    For position = 1 To Len(formulaBody)
        currentChar = Mid$(formulaBody, position, 1)

        Select Case currentChar
            Case QUOTE_CHAR
                'Toggle quotation state so commas inside quotes are ignored.
                'A doubled quote inside a quoted segment toggles the flag off
                'and straight back on, so the walker lands in the right state.
                'That is why an escaped quote needs no case of its own.
                insideQuotes = Not insideQuotes
            Case "("
                'Track depth of nested brackets to keep expressions intact
                If Not insideQuotes Then depth = depth + 1
            Case ")"
                'Close previously opened parenthesis when outside quotes
                If Not insideQuotes And depth > 0 Then depth = depth - 1
            Case TOKEN_SEPARATOR
                If Not insideQuotes And depth = 0 Then
                    'We reached a top-level comma: capture the segment
                    segmentText = Trim$(Mid$(formulaBody, segmentStart, position - segmentStart))
                    If LenB(segmentText) > 0 Then segments.Push segmentText
                    segmentStart = position + 1
                End If
        End Select
    Next

    'Push trailing segment after the final comma (or the lone element)
    segmentText = Trim$(Mid$(formulaBody, segmentStart, Len(formulaBody) - segmentStart + 1))
    If LenB(segmentText) > 0 Then segments.Push segmentText

    Set SplitFormulaSegments = segments
End Function
