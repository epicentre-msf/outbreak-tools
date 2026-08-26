VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Convert formulas from setup to Excel"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Convert formulas from setup to Excel")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, ReadOnlyPropertyAssignment, UseMeaningfulName, HungarianNotation

Option Explicit

'@class Formulas
'@description
'Reads a setup expression and answers the Excel formula it stands for. The
'expression is split into tokens, the variable names are resolved through an
'LLdictionary, and the custom functions are handed to CaseWhen, ChoiceFormula
'and ValueOfFormula. SetupFormula takes the expression in, Valid says whether
'it can be read, Reason names what stopped it, Varlists lists the variables it
'names and HasSetupVariables says whether it names any. ParsedLinelistFormula
'answers the formula for a data entry sheet, and ParsedAnalysisFormula the one
'for an analysis sheet.
'
'GROUPED FORMULAS
'-------------------------------------------------------------------------------
'IsGrouped says whether the expression uses one of the grouped custom
'functions. Such an expression is written out as the Excel aggregation
'FormulaData maps the token to.
'
'WHAT IS REPORTED
'-------------------------------------------------------------------------------
'Entries filed while an expression is read leave through HasChecking and
'CheckingValues.
'@depends BetterArray, Checking, LLdictionary, LLVariables, LLSheets, CaseWhen, ChoiceFormula, ValueOfFormula, FormulaData, FormulaCondition
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Stores cached parsing results, lazily resolved dependencies, and evaluation diagnostics.

Private Type TGroupContext
    Enabled As Boolean                          'Indicates whether the current formula targets grouped evaluation
    FunctionName As String                      'Root function detected in the setup expression
    AggregatorName As String                    'Excel aggregator applied during grouped evaluation
    UsesNative As Boolean                       'True when native *IFS function should be emitted
    CriteriaVariable As String                  'First argument (criteria range variable)
    ConditionVariable As String                 'Second argument (comparison variable)
    ResultVariable As String                    'Third argument (aggregation variable)
    TableName As String                         'Shared table name for the criteria and result variables
End Type

Private Type TFormulas
    dict As LLdictionary                       'Dictionary describing available variables and metadata
    data As FormulaData                        'Configuration of accepted Excel formulas and tokens
    setupForm As String                         'Original pseudo-formula provided by setup
    tokens As BetterArray                       'Tokenised representation of the parsed formula
    variables As BetterArray                    'Ordered list of variables detected in the formula
    containsSetupVariables As Boolean           'Indicates the presence of literals in the parsed expression
    isValid As Boolean                          'Flag capturing the latest validation outcome
    invalidReason As String                     'Human-readable explanation when validation fails
    validityComputed As Boolean                 'Marks whether EnsureEvaluation already processed a context
    lastContext As String                       'Context the cached evaluation belongs to
    conversionFailure As String                 'Why a custom function converter gave back nothing
    varsHelper As LLVariables                   'Lazy wrapper resolving variable metadata
    sheetsHelper As LLSheets                    'Lazy wrapper resolving worksheet addresses
    internalChecks As Checking                  'Diagnostic entries collected while parsing
    checkCounter As Long                        'Key counter for the diagnostic entries
    hasChecks As Boolean                        'True once a diagnostic entry has been written
    sealed As Boolean                           'Set by the factory, closes the setters
    group As TGroupContext                      'Captured state describing grouped formula evaluation
End Type

Private Const CLASSNAME As String = "Formulas"
Private Const FORMULA_DEFAULT_REASON As String = "The formula seems correct"
Private Const FORMULA_EMPTY_MESSAGE As String = "The formula is empty. No formula were found"
Private Const FORMULA_ANALYSIS_SINGLE_VAR As String = "Analysis formula can not consist of only one variable, you should use aggregation function"
Private Const FORMULA_UNKNOWN_TOKEN As String = "Unknown token '%1' encountered while parsing"
Private Const FORMULA_PAREN_MISMATCH As String = "The formula contains unmatched parentheses"
Private Const FORMULA_NEGATIVE_PAREN As String = "Closing parenthesis detected before opening one"
Private Const FORMULA_GROUP_ARGS_MISSING As String = "Grouped formulas must provide exactly three variable arguments."
Private Const FORMULA_GROUP_TABLE_MISMATCH As String = "Grouped formulas require the first and third variables to belong to the same table."
Private Const FORMULA_GROUP_MISSING_AGGREGATOR As String = "Grouped formula '%1' is not registered with an output aggregator."
Private Const FORMULA_GROUP_INVALID_GENERIC As String = "Grouped formula '%1' must specify a valid aggregator after the GROUP prefix."
Private Const FORMULA_GROUP_UNKNOWN_AGGREGATOR As String = "Grouped formula '%1' targets aggregator '%2', which is not allowed."

'The two contexts the class really has. Callers pass "analysis", "linelist"
'or "simple", and every branch below asks one question: is this the analysis
'context. EvaluationContext maps the three names onto these two, so a caller
'that asks Valid("linelist") and then ParsedLinelistFormula() parses once.
Private Const CONTEXT_ANALYSIS As String = "analysis"
Private Const CONTEXT_LINELIST As String = "linelist"

Private Const QUOTE_CHAR As String = """"
Private Const QUOTED_EQUAL As String = """="""
Private Const QUOTED_LESS_EQUAL As String = """<="""
Private Const QUOTED_GREATER_EQUAL As String = """>="""
Private Const QUOTED_NOT_EQUAL As String = """<>"""


Private this As TFormulas

'@section Factory
'===============================================================================
'@description Provides the single entry point to instantiate the parser and initialise core state.

'@label:create
'@sub-title Instantiate a Formulas helper ready to parse the provided setup expression
'@details
'Validates that the required dictionary and formula data objects are
'present, then stores them in a new instance alongside the setup
'expression. No eager parsing occurs; the formula is parsed lazily on
'first access to Valid, Varlists, or any parsed formula property.
'@param dict LLdictionary. Variable metadata provider.
'@param formData FormulaData. Catalogue of approved Excel functions and separators.
'@param setupForm String. Pseudo-formula to parse and validate.
'@return Formulas. Initialised instance.
'@throws ProjectError.ObjectNotInitialized When dict is Nothing.
'@throws ProjectError.ObjectNotInitialized When formData is Nothing.
'@export
Public Function Create(ByVal dict As LLdictionary, _
                       ByVal formData As FormulaData, _
                       ByVal setupForm As String) As Formulas

    If dict Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Dictionary is not found"
    If formData Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Formula Data not found"

    Dim instance As Formulas
    Set instance = New Formulas
    instance.Initialise dict, formData, setupForm
    instance.Seal
    Set Create = instance
End Function

'@label:seal
'@sub-title Prevent further changes to setup-only properties
'@details
'Marks the instance as sealed so guarded setters raise when invoked after
'construction. Called by the factory before returning. The class holds a
'parse result, and a setter running behind it would leave tokens that
'belong to another expression.
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
'@description Exposes configured dependencies, cached evaluation results, and parsed Excel expressions.

'@label:dictionary
'@sub-title Retrieve the underlying dictionary used during parsing
'@details
'Returns the LLdictionary previously supplied during Create or
'Initialise. Used by downstream helpers to resolve variable metadata.
'@return LLdictionary. Dictionary backing metadata lookups.
Public Property Get Dictionary() As LLdictionary
    Set Dictionary = this.dict
End Property

'@label:dictionary-set
'@sub-title Assign a new dictionary and reset dependent caches
'@details
'Replaces the current dictionary and invalidates all lazily resolved
'dependencies. The factory seals the instance, so this raises after
'construction.
'@param dict LLdictionary. Variable metadata provider.
'@throws ProjectError.ObjectNotInitialized When dict is Nothing.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Public Property Set Dictionary(ByVal dict As LLdictionary)
    GuardNotSealed "Dictionary"

    If dict Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Dictionary is not found"
    Set this.dict = dict
    Set this.varsHelper = Nothing
    Set this.sheetsHelper = Nothing
    InvalidateCaches
End Property

'@label:data
'@sub-title Retrieve the formula configuration catalogue
'@details
'Returns the FormulaData describing available Excel tokens,
'operators, and special characters used during tokenisation.
'@return FormulaData. Configuration catalogue.
Public Property Get Data() As FormulaData
    Set Data = this.data
End Property

'@label:data-set
'@sub-title Update the formula configuration catalogue and invalidate caches
'@details
'Replaces the current formula data and invalidates all cached parsing
'results. The factory seals the instance, so this raises after
'construction.
'@param formData FormulaData. Catalogue of allowed operators and functions.
'@throws ProjectError.ObjectNotInitialized When formData is Nothing.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Public Property Set Data(ByVal formData As FormulaData)
    GuardNotSealed "Data"

    If formData Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Formula Data not found"
    Set this.data = formData
    InvalidateCaches
End Property

'@label:setup-formula
'@sub-title Retrieve the original setup pseudo-formula
'@details
'Returns the raw expression as captured from the setup sheet during
'Create or the last assignment.
'@return String. Raw expression.
Public Property Get SetupFormula() As String
    SetupFormula = this.setupForm
End Property

'@label:setup-formula-set
'@sub-title Store the setup pseudo-formula and reset computed caches
'@details
'Replaces the stored expression and invalidates all cached parsing
'results. The factory seals the instance, so this raises after
'construction.
'@param setupForm String. Pseudo-formula captured from setup.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Public Property Let SetupFormula(ByVal setupForm As String)
    GuardNotSealed "SetupFormula"

    this.setupForm = setupForm
    InvalidateCaches
End Property

'@label:has-setup-variables
'@sub-title Determine if the parsed formula contains literal values
'@details
'Returns True when numeric or text literals were encountered during
'tokenisation. The flag is set during EvaluateFormula.
'@return Boolean. True when literals are present.
'@export
Public Property Get HasSetupVariables() As Boolean
    HasSetupVariables = this.containsSetupVariables
End Property

'@label:is-grouped
'@sub-title Indicate whether the parsed formula targets grouped evaluation
'@details
'Forces evaluation in the linelist context, then inspects the captured
'group metadata. Returns "Yes" when grouped logic applies, "No" otherwise.
'@return String. "Yes" for grouped formulas, "No" otherwise.
'@export
Public Property Get IsGrouped() As String
    Call EnsureEvaluation(CONTEXT_LINELIST)

    If this.group.Enabled Then
        IsGrouped = "Yes"
    Else
        IsGrouped = "No"
    End If

End Property

'@label:reason
'@sub-title Retrieve the last validation message associated with the given context
'@details
'Returns a human-readable explanation of the current validity state.
'When the formula is valid the default success message is returned;
'otherwise the recorded invalidation reason is returned.
'EnsureEvaluation answers with the validity flag, so one test decides
'which of the two messages comes back.
'@param formulaType Optional String. Context (analysis/linelist/simple).
'@return String. Explanation of the current validity state.
'@export
Public Property Get Reason(Optional ByVal formulaType As String = CONTEXT_ANALYSIS) As String
    If EnsureEvaluation(formulaType) Then
        Reason = FORMULA_DEFAULT_REASON
    Else
        Reason = this.invalidReason
    End If
End Property

'@label:valid
'@sub-title Check whether the formula is valid for the supplied context
'@details
'Triggers a lazy evaluation for the requested context, then returns
'the cached validity flag.
'@param formulaType Optional String. Context (analysis/linelist/simple).
'@return Boolean. True when parsing succeeded.
'@export
Public Property Get Valid(Optional ByVal formulaType As String = CONTEXT_ANALYSIS) As Boolean
    Valid = EnsureEvaluation(formulaType)
End Property

'@label:varlists
'@sub-title Provide a clone of the variables detected in the expression
'@details
'Returns a BetterArray containing the variable names encountered
'during tokenisation. The result is a clone so callers cannot mutate the
'internal cache.
'@param formulaType Optional String. Context (analysis/linelist/simple).
'@return BetterArray. Variable names.
'@export
Public Property Get Varlists(Optional ByVal formulaType As String = CONTEXT_ANALYSIS) As BetterArray
    Dim result As BetterArray

    If EnsureEvaluation(formulaType) Then
        If Not this.variables Is Nothing Then
            Set Varlists = this.variables.Clone()
            Exit Property
        End If
    End If

    Set result = New BetterArray
    result.LowerBound = 1
    Set Varlists = result
End Property

'@label:parsed-linelist-formula
'@sub-title Build the Excel expression for linelist worksheets
'@details
'Walks the cached token list and writes the output string as it goes,
'replacing variable tokens with worksheet addresses or structured
'references depending on the useTableName flag. Grouped formulas are
'delegated to BuildGroupedFormula. Custom tokens MEAN and N are mapped to
'their Excel equivalents (AVERAGE, COUNT).
'The token list is read where it lies. CrossTableFormula reaches one
'instance from 47 call sites, and each of those used to copy the whole list.
'@param useTableName Optional Boolean. Toggles structured references.
'@param tablePrefix Optional String. Prefixed to structured references.
'@return String. Excel formula referencing linelist variables.
'@export
Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
    Dim idx As Long
    Dim current As String
    Dim vars As LLVariables
    Dim sheets As LLSheets
    Dim builder As String

    If Not EnsureEvaluation(CONTEXT_LINELIST) Then Exit Property

    If this.group.Enabled Then
        ParsedLinelistFormula = BuildGroupedFormula(useTableName, tablePrefix)
        Exit Property
    End If

    Set vars = VariablesProvider()
    Set sheets = SheetProvider()

    For idx = this.tokens.LowerBound To this.tokens.UpperBound
        current = CStr(this.tokens.Item(idx))

        If vars.Contains(current) Then
            If useTableName Then
                builder = builder & tablePrefix & vars.Value(colName:="table name", varName:=current) & "[" & current & "]"
            Else
                builder = builder & sheets.VariableAddress(current)
            End If
        ElseIf StrComp(current, "MEAN", vbTextCompare) = 0 Then
            builder = builder & "AVERAGE"
        ElseIf StrComp(current, "N", vbTextCompare) = 0 Then
            builder = builder & "COUNT"
        Else
            builder = builder & current
        End If
    Next idx

    ParsedLinelistFormula = builder
End Property

'@label:parsed-analysis-formula
'@sub-title Build the Excel expression for aggregated analysis results
'@details
'Walks the cached token list and writes the output string as it goes,
'replacing variable tokens with conditional aggregation fragments supplied
'by formCond. Custom formula tokens are delegated to ParsedCustomFormula.
'Grouped formulas are delegated to BuildGroupedFormula.
'A custom token that expands to a whole call swallows the "(" and ")"
'tokens behind it, which is what the token removal used to do by shortening
'the list while the loop ran.
'@param formCond FormulaCondition. Provides conditional aggregation logic.
'@param tablePrefix Optional String. Prefix for table references.
'@param Connector Optional String. Connector between conditional fragments.
'@return String. Excel expression ready for analysis worksheets.
'@export
Public Property Get ParsedAnalysisFormula(ByVal formCond As FormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
    Dim vars As LLVariables
    Dim idx As Long
    Dim upperBound As Long
    Dim chunk As String
    Dim tableName As String
    Dim parsedCustom As String
    Dim builder As String

    If Not EnsureEvaluation(CONTEXT_ANALYSIS) Then Exit Property

    If this.group.Enabled Then
        ParsedAnalysisFormula = BuildGroupedFormula(True, tablePrefix)
        Exit Property
    End If

    If formCond Is Nothing Then Exit Property

    Set vars = VariablesProvider()

    'Every formula arrives with a fresh condition object. Handing it the
    'helper this instance already holds keeps each formula from reading the
    'dictionary sheet again for the same table.
    formCond.AdoptVariablesHelper vars, this.dict

    idx = this.tokens.LowerBound
    upperBound = this.tokens.UpperBound

    Do While idx <= upperBound
        chunk = CStr(this.tokens.Item(idx))

        If vars.Contains(chunk) Then
            tableName = vars.Value(varName:=chunk, colName:="table name")
            If Not formCond.Valid(this.dict, tableName) Then Exit Property
            builder = builder & formCond.ConditionString(varName:=chunk, _
                                                         TableName:=tablePrefix & tableName, _
                                                         Connector:=Connector)
        ElseIf IsCustomFormula(chunk) Then
            parsedCustom = ParsedCustomFormula(customFunction:=chunk, _
                                               formCond:=formCond, _
                                               tablePrefix:=tablePrefix)
            If LenB(parsedCustom) = 0 Then Exit Property
            builder = builder & parsedCustom
            If EmptyInvocationFollows(idx) Then idx = idx + 2
        Else
            builder = builder & chunk
        End If

        idx = idx + 1
    Loop

    ParsedAnalysisFormula = builder
End Property

'@section PrivateHelpers
'===============================================================================
'@description Internal utilities for dependency management, cache invalidation, and token resolution.

'@label:initialise
'@sub-title Store incoming dependencies and prepare caches
'@details
'Assigns the dictionary, formula data, and setup expression, then drops the
'lazily built helpers and invalidates the parse cache once.
'It stays Public because the factory calls it on a second instance, and it
'is guarded: after the factory seals that instance, this raises.
'@param dict LLdictionary. Variable metadata provider.
'@param formData FormulaData. Catalogue of allowed tokens and operators.
'@param setupForm String. Pseudo-formula to parse.
'@throws ProjectError.SomethingWentWrong When the instance is sealed.
Public Sub Initialise(ByVal dict As LLdictionary, _
                       ByVal formData As FormulaData, _
                       ByVal setupForm As String)
    GuardNotSealed "Initialise"

    Set this.dict = dict
    Set this.data = formData
    this.setupForm = setupForm
    Set this.varsHelper = Nothing
    Set this.sheetsHelper = Nothing
    InvalidateCaches
End Sub

'@label:is-custom-formula
'@sub-title Determine whether a token is one of the custom analysis formulas
'@details
'MEAN, N and N() are the whole set and they are compiled in. The match is
'case-sensitive, which is what the old list lookup did.
'AppendToken accepts these in the analysis context only, and
'ParsedAnalysisFormula tests the same three. Both tests are needed.
'@param token String. Candidate token.
'@return Boolean. True when the token is a custom analysis formula.
Private Function IsCustomFormula(ByVal token As String) As Boolean
    Select Case token
        Case "MEAN", "N", "N()"
            IsCustomFormula = True
    End Select
End Function

'@label:invalidate-caches
'@sub-title Reset cached parsing results and computed flags
'@details
'Clears the token list, variable list, validity flags, and group
'context so the next access triggers a full reparse from the stored
'setup expression.
Private Sub InvalidateCaches()
    Set this.tokens = Nothing
    Set this.variables = Nothing
    this.containsSetupVariables = False
    this.isValid = False
    this.invalidReason = vbNullString
    this.conversionFailure = vbNullString
    this.validityComputed = False
    this.lastContext = vbNullString
    ResetGroupContext
End Sub

'@label:reset-group-context
'@sub-title Clear any grouped-formula metadata captured during previous evaluations
'@details
'Resets all fields of the TGroupContext UDT to their default values.
Private Sub ResetGroupContext()
    With this.group
        .Enabled = False
        .FunctionName = vbNullString
        .AggregatorName = vbNullString
        .UsesNative = False
        .CriteriaVariable = vbNullString
        .ConditionVariable = vbNullString
        .ResultVariable = vbNullString
        .TableName = vbNullString
    End With
End Sub

'@label:variables-provider
'@sub-title Lazily instantiate the LLVariables helper from the dictionary
'@details
'Creates the LLVariables wrapper on first access and caches it for
'subsequent calls.
'@return LLVariables. Bound to the current dictionary.
Private Function VariablesProvider() As LLVariables
    If this.varsHelper Is Nothing Then
        Set this.varsHelper = LLVariables.Create(this.dict)
    End If
    Set VariablesProvider = this.varsHelper
End Function

'@label:sheet-provider
'@sub-title Lazily instantiate the LLSheets helper for address resolution
'@details
'Creates the LLSheets wrapper on first access and caches it for
'subsequent calls.
'@return LLSheets. Bound to the current dictionary.
Private Function SheetProvider() As LLSheets
    If this.sheetsHelper Is Nothing Then
        Set this.sheetsHelper = LLSheets.Create(this.dict)
    End If
    Set SheetProvider = this.sheetsHelper
End Function

'@label:evaluation-context
'@sub-title Map the three context names onto the two real contexts
'@details
'"analysis" is one context and everything else is the other. Production
'passes "linelist", the class passes itself "simple", and the public
'defaults are "analysis". Every branch of EvaluateFormula asks the same
'question, so holding three names made the same expression parse twice for
'the same answer.
'@param formulaType String. Context name as the caller wrote it.
'@return String. CONTEXT_ANALYSIS or CONTEXT_LINELIST.
Private Function EvaluationContext(ByVal formulaType As String) As String
    If StrComp(formulaType, CONTEXT_ANALYSIS, vbTextCompare) = 0 Then
        EvaluationContext = CONTEXT_ANALYSIS
    Else
        EvaluationContext = CONTEXT_LINELIST
    End If
End Function

'@label:ensure-evaluation
'@sub-title Guarantee that a parsing/validation run exists for the requested context
'@details
'Checks whether the cached evaluation matches the requested context.
'When a mismatch is detected or no evaluation has been performed, a
'full EvaluateFormula pass is triggered.
'@param formulaType String. Context (analysis/linelist/simple).
'@return Boolean. True when the cached evaluation is valid.
Private Function EnsureEvaluation(ByVal formulaType As String) As Boolean
    Dim context As String

    context = EvaluationContext(formulaType)

    If Not this.validityComputed Or StrComp(this.lastContext, context, vbBinaryCompare) <> 0 Then
        EvaluateFormula context
    End If

    EnsureEvaluation = this.isValid
End Function

'@label:evaluate-formula
'@sub-title Parse the setup expression, populate tokens, and capture validation results
'@details
'Converts the stored setup formula through custom function handlers,
'then tokenises the result while validating parentheses and literals.
'Short-circuits for single-variable or single-custom-function
'expressions. Delegates grouped formula detection to
'GroupContexExtractionSucceed.
'@param context String. CONTEXT_ANALYSIS or CONTEXT_LINELIST.
Private Sub EvaluateFormula(ByVal context As String)
    Dim converted As String
    Dim vars As LLVariables
    Dim dataSource As FormulaData
    Dim tokens As BetterArray
    Dim variableTokens As BetterArray
    Dim formulaHasSetupVariables As Boolean
    Dim reasonText As String

    InvalidateCaches
    this.lastContext = context

    converted = ConvertedSetupFormula()
    converted = CleanString(converted)

    Set vars = VariablesProvider()
    Set dataSource = this.data

    If LenB(converted) = 0 Then
        'A custom function that could not be converted also comes back empty.
        'Reporting that as "the formula is empty" told the user nothing about
        'an expression they did write.
        If LenB(this.conversionFailure) > 0 Then
            reasonText = this.conversionFailure
        Else
            reasonText = FORMULA_EMPTY_MESSAGE
        End If
        SetInvalid reasonText
        Exit Sub
    End If

    If vars.Contains(converted) Then
        If StrComp(context, CONTEXT_ANALYSIS, vbTextCompare) = 0 Then
            SetInvalid FORMULA_ANALYSIS_SINGLE_VAR
            Exit Sub
        Else
            Set tokens = New BetterArray
            tokens.LowerBound = 1
            tokens.Push converted

            Set variableTokens = New BetterArray
            variableTokens.LowerBound = 1
            variableTokens.Push converted

            this.containsSetupVariables = True 'has setup variables here means there are setup variables
            Set this.tokens = tokens
            Set this.variables = variableTokens
            this.isValid = True
            this.validityComputed = True
            Exit Sub
        End If
    End If

    If StrComp(context, CONTEXT_ANALYSIS, vbTextCompare) = 0 And IsCustomFormula(converted) Then
        Set tokens = New BetterArray
        tokens.LowerBound = 1
        tokens.Push converted

        Set this.tokens = tokens
        Set this.variables = New BetterArray
        this.variables.LowerBound = 1
        this.containsSetupVariables = True
        this.isValid = True
        this.validityComputed = True
        Exit Sub
    End If

    Set tokens = New BetterArray
    tokens.LowerBound = 1
    Set variableTokens = New BetterArray
    variableTokens.LowerBound = 1

    If Not TokeniseFormula(converted, context, tokens, variableTokens, formulaHasSetupVariables, reasonText) Then
        SetInvalid reasonText
        Exit Sub
    End If

    If Not GroupContexExtractionSucceed(tokens, variableTokens, vars, dataSource) Then Exit Sub

    this.containsSetupVariables = formulaHasSetupVariables
    Set this.tokens = tokens
    Set this.variables = variableTokens
    this.isValid = True
    this.validityComputed = True
End Sub

'@section GroupFormula
'===============================================================================
'@description Detects grouped formulas, validates arguments, and prepares metadata for downstream builders.

'@label:group-context-extraction
'@sub-title Analyse the tokenised expression and record grouped metadata when relevant
'@details
'Inspects the root function name to determine whether the formula
'targets grouped evaluation. When grouped, validates that exactly
'three variable arguments are present and that the first and third
'variables belong to the same table. Records the captured metadata
'in the TGroupContext UDT for BuildGroupedFormula.
'@param tokens BetterArray. Tokenised formula.
'@param variableTokens BetterArray. Variable names in encounter order.
'@param vars LLVariables. Helper exposing dictionary metadata.
'@param dataSource FormulaData. Configuration describing grouped functions.
'@return Boolean. True when grouped context is captured or not required.
Private Function GroupContexExtractionSucceed(ByVal tokens As BetterArray, _
                                       ByVal variableTokens As BetterArray, _
                                       ByVal vars As LLVariables, _
                                       ByVal dataSource As FormulaData) As Boolean
    Dim functionName As String
    Dim aggregatorName As String
    Dim usesNative As Boolean
    Dim isGeneric As Boolean
    Dim var1 As String
    Dim var2 As String
    Dim var3 As String
    Dim table1 As String
    Dim table3 As String
    Dim lowerBound As Long
    Dim reasonText As String

    ResetGroupContext
    GroupContexExtractionSucceed = False

    If tokens Is Nothing Then
        GroupContexExtractionSucceed = True
        Exit Function
    End If

    functionName = RootFunctionName(tokens)
    If LenB(functionName) = 0 Then
        GroupContexExtractionSucceed = True
        Exit Function
    End If

    isGeneric = IsGenericGroupFunction(functionName)

    If Not dataSource.IsGroupFunction(functionName) And Not isGeneric Then
        GroupContexExtractionSucceed = True
        Exit Function
    End If

    If isGeneric Then
        aggregatorName = ExtractGenericAggregator(functionName)
        If LenB(aggregatorName) = 0 Then
            reasonText = Replace(FORMULA_GROUP_INVALID_GENERIC, "%1", functionName, 1, 1, vbTextCompare)
            SetInvalid reasonText
            Exit Function
        End If
        If Not dataSource.ExcelFormulasIncludes(aggregatorName) Then
            reasonText = Replace(FORMULA_GROUP_UNKNOWN_AGGREGATOR, "%1", functionName, 1, 1, vbTextCompare)
            reasonText = Replace(reasonText, "%2", aggregatorName, 1, 1, vbTextCompare)
            SetInvalid reasonText
            Exit Function
        End If
        usesNative = False
    Else
        aggregatorName = dataSource.GroupAggregator(functionName)
        If LenB(aggregatorName) = 0 Then
            reasonText = Replace(FORMULA_GROUP_MISSING_AGGREGATOR, "%1", functionName, 1, 1, vbTextCompare)
            SetInvalid reasonText
            Exit Function
        End If
        usesNative = dataSource.GroupUsesNativeFunction(functionName)
    End If

    If variableTokens Is Nothing Then
        SetInvalid FORMULA_GROUP_ARGS_MISSING
        Exit Function
    End If

    If variableTokens.Length <> 3 Then
        SetInvalid FORMULA_GROUP_ARGS_MISSING
        Exit Function
    End If

    lowerBound = variableTokens.LowerBound
    var1 = CStr(variableTokens.Item(lowerBound))
    var2 = CStr(variableTokens.Item(lowerBound + 1))
    var3 = CStr(variableTokens.Item(lowerBound + 2))

    table1 = CStr(vars.Value(varName:=var1, colName:="table name"))
    table3 = CStr(vars.Value(varName:=var3, colName:="table name"))

    If LenB(table1) = 0 Or LenB(table3) = 0 Then
        SetInvalid FORMULA_GROUP_TABLE_MISMATCH
        Exit Function
    End If

    If StrComp(table1, table3, vbTextCompare) <> 0 Then
        SetInvalid FORMULA_GROUP_TABLE_MISMATCH
        Exit Function
    End If

    GroupContexExtractionSucceed = True
    With this.group
        .Enabled = True
        .FunctionName = functionName
        .AggregatorName = aggregatorName
        .UsesNative = usesNative
        .CriteriaVariable = var1
        .ConditionVariable = var2
        .ResultVariable = var3
        .TableName = table1
    End With
End Function

'@label:root-function-name
'@sub-title Retrieve the entry token corresponding to the root function
'@details
'Returns the first token from the tokenised expression, which is
'expected to be the function name when the formula is function-based.
'@param tokens BetterArray. Tokenised expression.
'@return String. Root token or empty string when unavailable.
Private Function RootFunctionName(ByVal tokens As BetterArray) As String
    If tokens Is Nothing Then Exit Function
    If tokens.Length = 0 Then Exit Function
    RootFunctionName = CStr(tokens.Item(tokens.LowerBound))
End Function

'@label:is-generic-group-function
'@sub-title Determine whether a function uses the GROUP prefix for grouped evaluation
'@details
'Returns True when the token starts with "GROUP" (case-insensitive). A
'token such as GROUPING_X is read as generic and its aggregator fails the
'whitelist check with a clear message, which is the wanted behaviour.
'@param functionName String. Token under inspection.
'@return Boolean. True when the token starts with "GROUP".
Private Function IsGenericGroupFunction(ByVal functionName As String) As Boolean
    If LenB(functionName) = 0 Then Exit Function
    IsGenericGroupFunction = (InStr(1, functionName, "GROUP", vbTextCompare) = 1)
End Function

'@label:extract-generic-aggregator
'@sub-title Parse the aggregator portion from a GROUP-prefixed function
'@details
'Splits the function name at the underscore separator and returns the
'trailing portion as the aggregator token. When no underscore is
'present, strips the "GROUP" prefix directly.
'@param functionName String. Grouped token (e.g. GROUP_SUM or GROUPS_SUM).
'@return String. Aggregator token to apply when building the Excel formula.
Private Function ExtractGenericAggregator(ByVal functionName As String) As String
    Dim normalized As String
    Dim separatorPosition As Long
    Dim candidate As String

    normalized = UCase$(functionName)
    separatorPosition = InStr(1, normalized, "_", vbBinaryCompare)

    If separatorPosition > 0 Then
        candidate = Mid$(normalized, separatorPosition + 1)
    Else
        candidate = Mid$(normalized, Len("GROUP") + 1)
    End If

    ExtractGenericAggregator = Trim$(candidate)
End Function

'@label:build-grouped-formula
'@sub-title Construct the grouped Excel expression using captured metadata
'@details
'Uses the TGroupContext UDT to resolve variable ranges and emit the
'appropriate aggregation expression. Native *IFS functions produce
'direct SUMIFS/COUNTIFS calls; non-native functions wrap the result
'in an IF-based array formula.
'@param useTableName Boolean. True to emit structured references.
'@param tablePrefix String. Prefix applied to structured references.
'@return String. Grouped Excel formula ready for consumption.
Private Function BuildGroupedFormula(ByVal useTableName As Boolean, _
                                     ByVal tablePrefix As String) As String
    Dim sheets As LLSheets
    Dim criteriaRange As String
    Dim conditionValue As String
    Dim resultRange As String
    Dim aggregator As String

    If Not this.group.Enabled Then Exit Function

    Set sheets = SheetProvider()

    aggregator = this.group.AggregatorName
    resultRange = GroupedDataRange(this.group.ResultVariable, useTableName, tablePrefix, sheets)
    criteriaRange = GroupedDataRange(this.group.CriteriaVariable, useTableName, tablePrefix, sheets)
    conditionValue = sheets.VariableAddress(this.group.ConditionVariable)

    If this.group.UsesNative Then
        Select Case aggregator
            Case "SUMIFS"
                BuildGroupedFormula = "SUMIFS(" & resultRange & ", " & criteriaRange & ", " & conditionValue & ")"
            Case "COUNTIFS"
                BuildGroupedFormula = "COUNTIFS(" & criteriaRange & ", " & conditionValue & ", " & resultRange & ", " & QUOTED_NOT_EQUAL & ")"
            Case Else
                BuildGroupedFormula = aggregator & "(" & resultRange & ", " & criteriaRange & ", " & conditionValue & ")"
        End Select
    Else
        BuildGroupedFormula = aggregator & "(IF(" & criteriaRange & "=" & conditionValue & ", " & resultRange & "))"
    End If
End Function

'@label:grouped-data-range
'@sub-title Resolve the appropriate range reference for a grouped variable
'@details
'Emits either a structured table reference or a sheet-level address
'depending on the useTableName flag.
'@param variableName String. Variable identifier.
'@param useTableName Boolean. True to emit structured references.
'@param tablePrefix String. Prefix applied to structured references.
'@param sheets LLSheets. Provider for address resolution.
'@return String. Excel reference targeting the grouped variable.
Private Function GroupedDataRange(ByVal variableName As String, _
                                  ByVal useTableName As Boolean, _
                                  ByVal tablePrefix As String, _
                                  ByVal sheets As LLSheets) As String
    Dim tableName As String

    tableName = this.group.TableName

    If useTableName Then
        GroupedDataRange = tablePrefix & tableName & "[" & variableName & "]"
    Else
        GroupedDataRange = sheets.VariableAddress(variableName)
    End If
End Function

'@section Tokenisation
'===============================================================================
'@description Breaks expressions into tokens, validates each chunk, and handles custom formula expansion.

'@label:tokenise-formula
'@sub-title Break the converted formula into tokens while validating parentheses and literals
'@details
'Walks the formula character by character, splitting on special
'characters defined by the FormulaData configuration. Tracks
'parenthesis depth and quotation state. Each extracted chunk is
'validated through AppendToken.
'The walk reads one character at a time. CaseWhen, ChoiceFormula and
'ValueOfFormula hold the same shape of walk over their own bodies. What
'made this loop expensive was the three Excel calls inside CleanString and
'the array copy inside the character lookup, and both of those are gone.
'@param formula String. Cleaned setup expression.
'@param context String. CONTEXT_ANALYSIS or CONTEXT_LINELIST.
'@param tokens BetterArray. Recipient for all tokens.
'@param variableTokens BetterArray. Recipient for variable-only tokens.
'@param HasSetupVariables Boolean. Flag capturing literal presence.
'@param failureReason String. Explanation when tokenisation fails.
'@return Boolean. True when parsing succeeds.
Private Function TokeniseFormula(ByVal formula As String, _
                                 ByVal context As String, _
                                 ByVal tokens As BetterArray, _
                                 ByVal variableTokens As BetterArray, _
                                 ByRef HasSetupVariables As Boolean, _
                                 ByRef failureReason As String) As Boolean
    Dim idx As Long
    Dim previousBreak As Long
    Dim currentChar As String
    Dim chunk As String
    Dim parenthesisLevel As Long
    Dim inQuotes As Boolean
    Dim dataSource As FormulaData
    Dim vars As LLVariables

    previousBreak = 1
    Set dataSource = this.data
    Set vars = VariablesProvider()

    For idx = 1 To Len(formula)
        currentChar = Mid$(formula, idx, 1)

        If currentChar = QUOTE_CHAR Then
            inQuotes = Not inQuotes
        ElseIf Not inQuotes Then
            Select Case currentChar
                Case "("
                    parenthesisLevel = parenthesisLevel + 1
                Case ")"
                    parenthesisLevel = parenthesisLevel - 1
                    If parenthesisLevel < 0 Then
                        failureReason = FORMULA_NEGATIVE_PAREN
                        Exit Function
                    End If
            End Select
        End If

        If Not inQuotes And dataSource.SpecialCharacterIncludes(currentChar) Then
            chunk = CleanString(Mid$(formula, previousBreak, idx - previousBreak))
            If Not AppendToken(chunk, context, tokens, variableTokens, vars, dataSource, HasSetupVariables, failureReason) Then Exit Function
            tokens.Push currentChar
            previousBreak = idx + 1
        End If
    Next idx

    chunk = CleanString(Mid$(formula, previousBreak, Len(formula) - previousBreak + 1))
    If Not AppendToken(chunk, context, tokens, variableTokens, vars, dataSource, HasSetupVariables, failureReason) Then Exit Function

    If parenthesisLevel <> 0 Then
        failureReason = FORMULA_PAREN_MISMATCH
        Exit Function
    End If

    TokeniseFormula = True
End Function

'@label:append-token
'@sub-title Validate and append the provided chunk to the tokens collection
'@details
'Classifies the chunk as a grouped function, known variable, boolean
'literal, registered Excel formula, custom formula, numeric literal,
'or quoted string. Unknown chunks cause a validation failure.
'The grouped-function test runs on the FIRST token only. That is what
'stops SUM(SUMIFS(...)) from being read as a grouped formula.
'@param chunk String. Candidate token.
'@param context String. CONTEXT_ANALYSIS or CONTEXT_LINELIST.
'@param tokens BetterArray. Recipient for all tokens.
'@param variableTokens BetterArray. Recipient for variable-only tokens.
'@param vars LLVariables. Metadata provider.
'@param dataSource FormulaData. Formula metadata provider.
'@param HasSetupVariables Boolean. Flag capturing literal presence.
'@param failureReason String. Explanation when parsing fails.
'@return Boolean. True when the chunk is accepted.
Private Function AppendToken(ByVal chunk As String, _
                             ByRef context As String, _
                             ByRef tokens As BetterArray, _
                             ByRef variableTokens As BetterArray, _
                             ByRef vars As LLVariables, _
                             ByRef dataSource As FormulaData, _
                             ByRef HasSetupVariables As Boolean, _
                             ByRef failureReason As String) As Boolean
    Dim upperChunk As String

    If LenB(chunk) = 0 Then
        AppendToken = True
        Exit Function
    End If

    upperChunk = UCase$(chunk)

    'Grouped formula
    If (tokens.Length = 0) Then
        If IsGenericGroupFunction(upperChunk) Then
            tokens.Push chunk
            AppendToken = True
            Exit Function
        End If

        If dataSource.IsGroupFunction(upperChunk) Then
            tokens.Push chunk
            AppendToken = True
            Exit Function
        End If
    End If

    If vars.Contains(chunk) Then
        variableTokens.Push chunk
        tokens.Push chunk
        HasSetupVariables = True
        AppendToken = True
        Exit Function
    End If

    If StrComp(upperChunk, "TRUE", vbTextCompare) = 0 Or _
       StrComp(upperChunk, "FALSE", vbTextCompare) = 0 Then
        tokens.Push upperChunk
        AppendToken = True
        Exit Function
    End If

    If dataSource.ExcelFormulasIncludes(upperChunk) Then
        tokens.Push upperChunk
        AppendToken = True
        Exit Function
    End If

    If StrComp(context, CONTEXT_ANALYSIS, vbTextCompare) = 0 And IsCustomFormula(chunk) Then
        tokens.Push chunk
        AppendToken = True
        Exit Function
    End If

    If IsNumericLiteral(chunk) Then
        tokens.Push chunk
        AppendToken = True
        Exit Function
    End If

    If IsQuotedString(chunk) Then
        tokens.Push chunk
        AppendToken = True
        Exit Function
    End If

    failureReason = Replace(FORMULA_UNKNOWN_TOKEN, "%1", chunk)
    AppendToken = False
End Function

'@label:is-quoted-string
'@sub-title Determine whether the supplied value is wrapped in quotes
'@details
'Returns True when the string starts and ends with a double-quote
'character and has a length of at least two.
'@param value String. Candidate literal.
'@return Boolean. True when the string is quoted.
Private Function IsQuotedString(ByVal value As String) As Boolean
    If Len(value) < 2 Then Exit Function
    IsQuotedString = (Left$(value, 1) = QUOTE_CHAR And Right$(value, 1) = QUOTE_CHAR)
End Function

'@label:is-numeric-literal
'@sub-title Determine whether the supplied chunk is an English-form numeric literal
'@details
'Accepts digits around at most one dot. The convention keeps setup formulas
'in English, and Range.Formula reads US syntax on every host, so the English
'form is the one form a literal may take here. IsNumeric would read the
'chunk in the HOST's regional separators, and on a comma machine it refused
'every dot literal a setup file carried. A sign never reaches this test:
''+' and '-' are separator characters, so they split off before the chunk
'arrives.
'@param chunk String. Candidate literal.
'@return Boolean. True when the chunk holds at least one digit and nothing besides digits and a single dot.
Private Function IsNumericLiteral(ByVal chunk As String) As Boolean
    Dim idx As Long
    Dim oneChar As String
    Dim dotSeen As Boolean
    Dim digitSeen As Boolean

    For idx = 1 To Len(chunk)
        oneChar = Mid$(chunk, idx, 1)
        If oneChar = "." Then
            If dotSeen Then Exit Function
            dotSeen = True
        ElseIf oneChar >= "0" And oneChar <= "9" Then
            digitSeen = True
        Else
            Exit Function
        End If
    Next idx

    IsNumericLiteral = digitSeen
End Function

'@label:clean-string
'@sub-title Normalise the incoming token by removing control characters and duplicate spaces
'@details
'Replaces non-breaking spaces, drops control characters, squeezes runs of
'spaces to one, and trims the ends. This is what the worksheet SUBSTITUTE,
'CLEAN and TRIM trio did, written in plain VBA: the three of them crossed
'into Excel once each, and this runs once per token of every formula in the
'workbook.
'AscW answers a negative number above code point 32767, so the control-code
'test guards on zero. Worksheet TRIM squeezes runs of spaces INSIDE the
'text as well as at the ends, so that is reproduced.
'@param value String. Token to clean.
'@return String. Trimmed and cleaned representation.
Private Function CleanString(ByVal value As String) As String
    Dim source As String
    Dim result As String
    Dim counter As Long
    Dim charCode As Long
    Dim oneChar As String
    Dim previousWasSpace As Boolean

    'ChrW$(160) is U+00A0 NO-BREAK SPACE. ChrW$ reads 160 as a Unicode code
    'point, so the answer is the same on every machine. Chr$ would read it as a
    'byte position in the machine's ANSI codepage, which on this Mac holds the
    'dagger, and the non-breaking spaces would all survive. It cannot sit in a
    'Const either way, since a Const takes a constant expression.
    source = Replace$(value, ChrW$(160), " ")

    For counter = 1 To Len(source)
        oneChar = Mid$(source, counter, 1)
        charCode = AscW(oneChar)

        If charCode >= 0 And charCode < 32 Then
            'Dropped, the same as the worksheet CLEAN function.
        ElseIf oneChar = " " Then
            If Not previousWasSpace Then
                result = result & " "
                previousWasSpace = True
            End If
        Else
            result = result & oneChar
            previousWasSpace = False
        End If
    Next

    CleanString = Trim$(result)
End Function

'@label:parsed-custom-formula
'@sub-title Delegate parsing to custom functions when present in analysis context
'@details
'Handles MEAN (mapped to AVERAGE), N/N() (mapped to COUNTIFS), and
'passes through unknown custom identifiers unchanged. The COUNTIFS
'expansion builds criteria pairs from the FormulaCondition.
'The table always comes from the FormulaCondition. It used to be carried
'from whichever variable token came first, so the same expression gave two
'answers depending on the order of its tokens.
'@param customFunction String. Identifier of the custom formula.
'@param formCond FormulaCondition. Provides condition logic.
'@param tablePrefix Optional String. Prefix for structured references.
'@return String. Excel fragment produced by the custom handler.
Private Function ParsedCustomFormula(ByVal customFunction As String, _
                                     ByVal formCond As FormulaCondition, _
                                     Optional ByVal tablePrefix As String = vbNullString) As String
    Dim cond As BetterArray
    Dim condVars As BetterArray
    Dim builder As String
    Dim customTableName As String
    Dim idx As Long

    customTableName = tablePrefix & formCond.VariablesTable(this.dict)

    Select Case UCase$(customFunction)
        Case "N", "N()"
            Set cond = formCond.Conditions()
            Set condVars = formCond.Variables()
            builder = customTableName & "[" & condVars.Item(condVars.LowerBound) & "]" & ", " & ClearCountIf(cond.Item(condVars.LowerBound))
            For idx = condVars.LowerBound + 1 To condVars.UpperBound
                builder = builder & "," & customTableName & "[" & condVars.Item(idx) & "]" & "," & ClearCountIf(cond.Item(idx))
            Next idx
            builder = "COUNTIFS(" & builder & ")"
        Case "MEAN"
            builder = "AVERAGE"
        Case Else
            builder = customFunction
    End Select

    ParsedCustomFormula = builder
End Function

'@label:empty-invocation-follows
'@sub-title Report whether the two tokens after a custom formula are an empty ()
'@details
'When a custom formula token is replaced inline, the "(" and ")" tokens
'behind it become redundant and are skipped.
'@param functionIndex Long. Index of the custom function token.
'@return Boolean. True when the next two tokens are "(" then ")".
Private Function EmptyInvocationFollows(ByVal functionIndex As Long) As Boolean
    Dim upperBound As Long

    If this.tokens Is Nothing Then Exit Function

    upperBound = this.tokens.UpperBound
    If functionIndex + 2 > upperBound Then Exit Function

    If StrComp(CStr(this.tokens.Item(functionIndex + 1)), "(", vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(CStr(this.tokens.Item(functionIndex + 2)), ")", vbBinaryCompare) <> 0 Then Exit Function

    EmptyInvocationFollows = True
End Function

'@label:clear-count-if
'@sub-title Remove equality fragments from COUNTIF-style conditions
'@details
'Sanitises predicate fragments by stripping equality operators and
'normalising less-than and greater-than comparisons for COUNTIFS
'compatibility. FormulaCondition builds the fragments this rewrites, so
'the two have to keep agreeing on their shape.
'@param value String. Predicate fragment.
'@return String. Sanitised predicate suitable for COUNTIFS.
Private Function ClearCountIf(ByVal value As String) As String
    Dim cleared As String
    cleared = Replace(value, QUOTED_EQUAL & "&", vbNullString)
    cleared = Replace(cleared, "=", vbNullString)
    cleared = Replace(cleared, QUOTE_CHAR & QUOTE_CHAR & QUOTE_CHAR & QUOTE_CHAR, QUOTE_CHAR & QUOTE_CHAR)
    cleared = Replace(cleared, "<>", "######")
    cleared = Replace(cleared, "<", QUOTED_LESS_EQUAL & " & ")
    cleared = Replace(cleared, ">", QUOTED_GREATER_EQUAL & " & ")
    cleared = Replace(cleared, "######", "<>")
    ClearCountIf = cleared
End Function

'@label:converted-setup-formula
'@sub-title Convert higher-level setup directives into plain Excel expressions
'@details
'Detects CASE_WHEN, CHOICE_FORMULA, and VALUE_OF prefixes in the
'stored setup expression and delegates to the corresponding parser
'class. Returns the original expression when no prefix matches.
'The expression is trimmed first. The three prefix tests anchor at
'position 1, so a leading space sent the raw CASE_WHEN text to the
'tokeniser, which reported it as an unknown token.
'When a converter answers with an empty string, its own failure message is
'held so the caller can report what really went wrong.
'@return String. Converted formula ready for tokenising.
Private Function ConvertedSetupFormula() As String
    Dim result As String
    Dim caseWhenObj As CaseWhen
    Dim choiceObj As ChoiceFormula
    Dim valueObj As ValueOfFormula

    this.conversionFailure = vbNullString
    result = Trim$(this.setupForm)

    If InStr(1, result, "CASE_WHEN", vbTextCompare) = 1 Then
        Set caseWhenObj = CaseWhen.Create(result)
        result = caseWhenObj.ParsedFormula()
        If LenB(result) = 0 Then this.conversionFailure = caseWhenObj.FailureReason
    ElseIf InStr(1, result, "CHOICE_FORMULA", vbTextCompare) = 1 Then
        Set choiceObj = ChoiceFormula.Create(result)
        result = choiceObj.ParsedFormula()
        If LenB(result) = 0 Then this.conversionFailure = choiceObj.FailureReason
    ElseIf InStr(1, result, "VALUE_OF", vbTextCompare) = 1 Then
        Set valueObj = ValueOfFormula.Create(result, this.dict, VariablesProvider(), SheetProvider())
        result = valueObj.ConvertedFormula
        If LenB(result) = 0 Then this.conversionFailure = valueObj.FailureReason
    End If

    ConvertedSetupFormula = result
End Function

'@label:set-invalid
'@sub-title Record an invalidation reason and log the associated checking entry
'@details
'Stores the failure message, marks the formula as invalid, and logs
'the message through the internal Checking store.
'@param message String. Explanation of the validation failure.
'@param scope Optional Byte. Severity level persisted in Checking.
Private Sub SetInvalid(ByVal message As String, Optional ByVal scope As Byte = checkingError)
    this.invalidReason = message
    this.isValid = False
    this.validityComputed = True
    LogCheck message, scope
End Sub

'@section Checkings
'===============================================================================
'@description Aggregates informational and error traces emitted during parsing.

'@label:log-check
'@sub-title Record a diagnostic entry within the internal checking store
'@details
'Lazily creates the Checking instance on first use, then appends
'the labelled entry with the given severity scope.
'@param label String. Message to log.
'@param scope Optional Byte. Severity indicator.
Private Sub LogCheck(ByVal label As String, _
                     Optional ByVal scope As Byte = checkingNote)

    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create(CLASSNAME)
    End If

    this.internalChecks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:has-checkings
'@sub-title Indicate whether any diagnostic entries were recorded
'@details
'Returns the flag set by LogCheck. VarWriter reads this name at four call
'sites, so the spelling stays until a session owns that file.
'@return Boolean. True when checkings exist.
'@export
Public Property Get HasChecking() As Boolean
    HasChecking = this.hasChecks
End Property

'@label:checking-values
'@sub-title Expose the collected checking entries if available
'@details
'Returns the internal Checking store when at least one entry has
'been recorded. Returns Nothing otherwise.
'@return Checking. Logged diagnostics.
'@export
Public Property Get CheckingValues() As Checking
    If Not this.hasChecks Then Exit Property
    Set CheckingValues = this.internalChecks
End Property

'@section ErrorHandling
'===============================================================================
'@description Centralises the error raising pattern for consistent ProjectError usage.

'@label:throw-error
'@sub-title Raise a VBA error with the class signature
'@details
'Wrapper around Err.Raise that standardises the source to the class
'constant CLASSNAME for consistent stack traces. The message goes through
'as written, the same as the other five classes in this folder.
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub
