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

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Custom Functions")
'@ModuleDescription("Parse and expose CHOICE_FORMULA expressions")
'@depends CaseWhen, BetterArray
'@label:ChoiceFormula
'@section Overview
'@details Converts CHOICE_FORMULA custom functions into CASE_WHEN representations, exposing helper metadata for downstream validation.

Option Explicit
Implements IChoiceFormula

'@section State
'===============================================================================

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

'@label:Create
'@sub-title Instantiate a ChoiceFormula parser bound to a formula
'@details Trims the incoming formula and returns the interface for fluent operations.
'@param formula String representation of a CHOICE_FORMULA call.
'@return IChoiceFormula ready to be queried.
Public Function Create(ByVal formula As String) As IChoiceFormula
    With New ChoiceFormula
        .ChoiceSpecifications = formula
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================

'@label:Self
'@sub-title Expose the current instance through its interface
'@return IChoiceFormula reference to the current object.
Public Property Get Self() As IChoiceFormula
    Set Self = Me
End Property

'@label:ChoiceSpecifications
'@sub-title Retrieve the stored CHOICE_FORMULA expression
'@return String original formula specification.
Public Property Get ChoiceSpecifications() As String
    ChoiceSpecifications = this.formula
End Property

'@sub-title Store the CHOICE_FORMULA expression and reset caches
'@param formula String expression to parse.
Public Property Let ChoiceSpecifications(ByVal formula As String)
    this.formula = Trim$(formula)
    ResetCaches
End Property

'@section Validation
'===============================================================================

'@label:Valid
'@sub-title Determine whether the stored formula is a CHOICE_FORMULA with arguments
'@details Requires the CHOICE_FORMULA token, balanced parentheses, and at least one condition/result pair.
'@return Boolean True when the specification can be converted to CASE_WHEN.
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

'@label:ParsedFormula
'@sub-title Convert CHOICE_FORMULA into nested IF statements through CASE_WHEN
'@details Relies on the CaseWhen class to build the Excel-compatible representation.
'@return String expression ready for worksheet consumption.
Private Function parsedFormula() As String
    Dim caseObj As ICaseWhen

    If Not Valid() Then Exit Function

    Set caseObj = CaseWhenWrapper()
    If caseObj Is Nothing Then Exit Function

    parsedFormula = caseObj.parsedFormula()
End Function

'@label:Categories
'@sub-title Extract textual labels from the CHOICE_FORMULA results
'@details Delegates to the CaseWhen converter to reuse category parsing logic.
'@return BetterArray containing quoted category labels.
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

    Set categoriesTable = caseObj.Categories()
    Set Categories = categoriesTable.Clone()
End Function

'@label:ChoiceName
'@sub-title Retrieve the name of the choice list referenced by the formula
'@return String choice list identifier.
Private Function choiceName() As String
    choiceName = ResolvedChoiceName()
End Function

'@section PrivateHelpers
'===============================================================================

'@sub-title Reset cached artefacts when the formula changes
Private Sub ResetCaches()
    Set this.cachedArguments = Nothing
    Set this.cachedCaseWhen = Nothing
    this.cachedChoiceName = vbNullString
End Sub

'@sub-title Resolve the portion of the formula between parentheses
'@return String containing the comma-delimited arguments.
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

    working = Left$(working, Len(working) - 1)
    SanitisedFormula = Trim$(working)
End Function

'@sub-title Split the sanitised formula into individual arguments
'@return BetterArray containing the choice name and condition/result entries.
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

    segmentText = Trim$(Mid$(body, segmentStart, formulaLength - segmentStart + 1))
    If LenB(segmentText) > 0 Then arguments.Push segmentText

    Set this.cachedArguments = arguments.Clone()
    Set ChoiceArguments = this.cachedArguments.Clone()
End Function

'@sub-title Determine whether the arguments contain empty entries
'@param arguments BetterArray list of parsed arguments.
'@return Boolean True when an argument is blank.
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

'@sub-title Build the equivalent CASE_WHEN formula from parsed arguments
'@return String CASE_WHEN expression or empty when invalid.
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

    builder = builder & ")"
    CaseWhenFormula = builder
End Function

'@sub-title Provide a lazily created CaseWhen instance
'@return ICaseWhen representing the converted formula.
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

'@sub-title Resolve and cache the choice list name
'@return String identifier extracted from the first argument.
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

'@sub-title Raise a project-specific error
'@param errNumber ProjectError describing the failure.
'@param message String message to propagate.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub

'@section Interface
'===============================================================================

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
