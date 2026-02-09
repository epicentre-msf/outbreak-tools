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
Implements ICaseWhen

'@class CaseWhen
'@description
'Parses CASE_WHEN custom formulas written in setup expressions and rewrites
'them as nested Excel IF statements. The class splits the formula body into
'alternating condition/result pairs while respecting quoted segments and
'nested parentheses, then builds the equivalent nested IF expression. It
'also extracts textual category labels for downstream analysis consumers.
'@depends BetterArray
'@version 1.0 (2026-02-09)
'@author OutbreakTools


'@section State
'===============================================================================
'@description Backing UDT, constants, and cached parsing table.

Private Type TCaseWhen
    formula As String
    parsingTable As BetterArray
End Type

Private this As TCaseWhen

Private Const TOKEN_HEADER As String = "CASE_WHEN("
Private Const TOKEN_SEPARATOR As String = ","
Private Const QUOTE_CHAR As String = """"

'@section Factory
'===============================================================================

'@label:create
'@sub-title Instantiate a CaseWhen parser bound to a formula
'@details
'Trims the incoming formula, stores it in a new instance, and returns
'the ICaseWhen interface reference for fluent downstream consumption.
'No eager parsing occurs; the formula is parsed lazily on first access
'to Valid, ParsedFormula, or Categories.
'@param formula String. The CASE_WHEN expression to parse.
'@return ICaseWhen. A ready-to-query parser instance.
'@export
Public Function Create(ByVal formula As String) As ICaseWhen
    With New CaseWhen
        .casewhenFormula = formula
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================
'@description Properties that expose the internal state and interface pointer.

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to its interface. Used internally
'by the factory method to return an ICaseWhen reference.
'@return ICaseWhen. Reference to the current object.
Public Property Get Self() As ICaseWhen
    Set Self = Me
End Property

'@label:formula
'@prop-title Retrieve the raw CASE_WHEN formula
'@details
'Returns the original formula text as stored after trimming during
'assignment. This is the full CASE_WHEN(...) expression.
'@return String. Original formula text.
Public Property Get casewhenFormula() As String
    casewhenFormula = this.formula
End Property

'@label:formula-set
'@prop-title Store the CASE_WHEN formula and reset cached data
'@details
'Trims the incoming formula and stores it, then invalidates the
'cached parsing table so subsequent reads trigger a fresh parse.
'@param formula String. Expression to analyse.
Public Property Let casewhenFormula(ByVal formula As String)
    this.formula = Trim$(formula)
    ResetCache
End Property

'@section Validation
'===============================================================================
'@description Checks whether the stored formula matches the CASE_WHEN pattern.

'@label:valid
'@prop-title Determine whether the stored formula is a CASE_WHEN
'@details
'Accepts formulas that begin with the CASE_WHEN( token
'(case-insensitive) and contain content beyond the header. Returns
'False for empty strings or formulas that do not match the token.
'@return Boolean. True when the formula is parseable.
Private Property Get Valid() As Boolean
    Dim formula As String
    formula = Trim$(casewhenFormula)
    If LenB(formula) = 0 Then Exit Property
    Valid = StrComp(Left$(UCase$(formula), Len(TOKEN_HEADER)), TOKEN_HEADER, vbBinaryCompare) = 0
End Property

'@section ParsingInfrastructure
'===============================================================================
'@description Lazily computed and cached argument table.

'@label:case-when-table
'@prop-title Retrieve the cached list of CASE_WHEN arguments
'@details
'Splits the formula into alternating condition/result entries while
'respecting quoted segments and nested parentheses. The result is
'cached so subsequent accesses avoid reparsing the string. Returns
'a clone to prevent callers from mutating the cache.
'@return BetterArray. The ordered CASE_WHEN arguments.
Private Property Get CaseWhenTable() As BetterArray
    Dim parsingTable As BetterArray
    Dim cleanedFormula As String

    If Not (this.parsingTable Is Nothing) Then
        Set CaseWhenTable = this.parsingTable.Clone()
        Exit Property
    End If

    Set parsingTable = New BetterArray
    parsingTable.LowerBound = 1

    If Valid() Then
        cleanedFormula = SanitisedFormula()
        If LenB(cleanedFormula) > 0 Then
            Set parsingTable = SplitFormulaSegments(cleanedFormula)
        End If
    End If

    'Cache the computed table so subsequent calls avoid reparsing the string.
    Set this.parsingTable = parsingTable.Clone()
    Set CaseWhenTable = this.parsingTable.Clone()
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
'nested IF calls.
'@return String. Nested IF expression ready for worksheet consumption.
Private Function parsedFormula() As String

    Dim counter As Long
    Dim formula As String
    Dim parenthesis As String
    Dim parsingTable As BetterArray
    Dim totalItems As Long

    formula = vbNullString
    parenthesis = vbNullString
    counter = 1

    If Valid() Then
        Set parsingTable = CaseWhenTable()
        totalItems = parsingTable.Length
        If totalItems = 0 Then Exit Function

        'Convert to nested IFs. Excel supports up to 64 nested IF calls.
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
    End If
    parsedFormula = formula
End Function

'@label:categories
'@sub-title Extract textual result labels from the CASE_WHEN arguments
'@details
'Evaluates each result expression (even-indexed entries in the argument
'table) and returns the literal values that will appear to end-users.
'Quoted entries are stripped of their quotes; non-quoted entries are
'wrapped in quotes for consistent downstream handling.
'@return BetterArray. Collection containing quoted category labels.
Private Function Categories() As BetterArray
    Dim categoriesTable As BetterArray
    Dim counter As Long
    Dim parsingTable As BetterArray
    Dim levelValue As String
    Dim totalItems As Long

    Set categoriesTable = New BetterArray
    categoriesTable.LowerBound = 1

    If Valid() Then
        counter = 2
        Set parsingTable = CaseWhenTable()
        totalItems = parsingTable.Length

        Do While counter <= totalItems
            'Result expressions sit at even indexes (2,4,6...)
            If Left$(parsingTable.Item(counter), 1) = QUOTE_CHAR Then
                'Quoted characters represent literal labels
                levelValue = Replace$(parsingTable.Item(counter), QUOTE_CHAR, vbNullString)
            Else
                'Not Quoted character, the user should not use case_when in analysis
                levelValue = QUOTE_CHAR & parsingTable.Item(counter) & QUOTE_CHAR
            End If
            If LenB(levelValue) > 0 Then categoriesTable.Push levelValue
            If (counter = (totalItems - 1)) Then
                counter = counter + 1
            Else
                counter = counter + 2
            End If
        Loop
    End If

    Set Categories = categoriesTable.Clone()
End Function

'@section PrivateHelpers
'===============================================================================
'@description Internal utilities for cache management and formula segmentation.

'@label:reset-cache
'@sub-title Reset cached artefacts when the source formula changes
'@details
'Sets the cached parsing table to Nothing so the next access to
'CaseWhenTable triggers a fresh parse from the stored formula.
Private Sub ResetCache()
    Set this.parsingTable = Nothing
End Sub

'@label:sanitised-formula
'@sub-title Remove the CASE_WHEN header and closing parenthesis
'@details
'Strips the leading CASE_WHEN( token and the trailing closing
'parenthesis to expose the raw comma-delimited argument body.
'@return String. Inner content of the CASE_WHEN call.
Private Function SanitisedFormula() As String
    Dim trimmedFormula As String
    trimmedFormula = Trim$(casewhenFormula)
    If LenB(trimmedFormula) = 0 Then Exit Function

    'Remove the leading token and explicit trailing parenthesis
    trimmedFormula = Mid$(trimmedFormula, Len(TOKEN_HEADER) + 1)
    If Right$(trimmedFormula, 1) = ")" Then
        trimmedFormula = Left$(trimmedFormula, Len(trimmedFormula) - 1)
    End If

    SanitisedFormula = Trim$(trimmedFormula)
End Function

'@label:split-formula-segments
'@sub-title Split the CASE_WHEN body into trimmed argument segments
'@details
'Walks the formula body character by character, tracking quotation
'state and parenthesis depth. Splits on top-level commas only,
'preserving nested expressions and quoted strings intact.
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

    Set segments = New BetterArray
    segments.LowerBound = 1

    segmentStart = 1

    For position = 1 To Len(formulaBody)
        currentChar = Mid$(formulaBody, position, 1)

        Select Case currentChar
            Case QUOTE_CHAR
                'Toggle quotation state so commas inside quotes are ignored
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
    '@Ignore AssignmentNotUsed
    segmentText = Trim$(Mid$(formulaBody, segmentStart, Len(formulaBody) - segmentStart + 1))
    If LenB(segmentText) > 0 Then segments.Push segmentText

    Set SplitFormulaSegments = segments
End Function

'@section Interface
'===============================================================================
'@description Delegated members satisfying the ICaseWhen contract.

Private Property Get ICaseWhen_Valid() As Boolean
    ICaseWhen_Valid = Valid()
End Property

Private Property Get ICaseWhen_ParsedFormula() As String
    ICaseWhen_ParsedFormula = parsedFormula()
End Property

Private Property Get ICaseWhen_Categories() As BetterArray
    Set ICaseWhen_Categories = Categories()
End Property
