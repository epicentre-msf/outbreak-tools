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
'@depends BetterArray, Checking, LLdictionary, LLVariables, LLSheets, CaseWhen, ChoiceFormula, ValueOfFormula
'@section Overview
'===============================================================================
'@details Parses setup expressions, validates token sequences, and exposes Excel-compatible formulas for linelist and analysis contexts.

Option Explicit
Implements IFormulas

'@section State
'===============================================================================
'@details Stores cached parsing results, lazily resolved dependencies, and evaluation diagnostics.

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
    dict As ILLdictionary                       'Dictionary describing available variables and metadata
    data As IFormulaData                        'Configuration of accepted Excel formulas and tokens
    setupForm As String                         'Original pseudo-formula provided by setup
    tokens As BetterArray                       'Tokenised representation of the parsed formula
    variables As BetterArray                    'Ordered list of variables detected in the formula
    containsSetupVariables As Boolean           'Indicates the presence of literals in the parsed expression
    isValid As Boolean                          'Flag capturing the latest validation outcome
    invalidReason As String                     'Human-readable explanation when validation fails
    validityComputed As Boolean                 'Marks whether EnsureEvaluation already processed a context
    lastFormulaType As String                   'Context (analysis/simple) associated with the cached evaluation
    group As TGroupContext                      'Captured state describing grouped formula evaluation
End Type

Private Type TDependencies
    variables As ILLVariables                   'Lazy wrapper resolving variable metadata
    sheets As ILLSheets                         'Lazy wrapper resolving worksheet addresses
    customFormulas As BetterArray               'Registered custom formula identifiers
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

Private this As TFormulas
Private deps As TDependencies
Private internalChecks As IChecking
Private checkCounter As Long
Private hasChecks As Boolean

'@section Factory
'===============================================================================
'@details Provides the single entry point to instantiate the parser and initialise core state.

'@label:Create
'@jump:Create
'@sub-title Instantiate a Formulas helper ready to parse the provided setup expression
'@param dict ILLdictionary supplying variable metadata.
'@param formData IFormulaData listing approved Excel functions and separators.
'@param setupForm String pseudo-formula to parse and validate.
'@return IFormulas initialised instance.
'@export
Public Function Create(ByVal dict As ILLdictionary, _
                       ByVal formData As IFormulaData, _
                       ByVal setupForm As String) As IFormulas

    If dict Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Dictionary is not found"
    If formData Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Formula Data not found"

    With New Formulas
        .Initialise dict, formData, setupForm
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================
'@details Exposes configured dependencies, cached evaluation results, and parsed Excel expressions.

'@label:Self
'@sub-title Return the interface pointer for the current instance
'@return IFormulas reference representing the current object.
Public Property Get Self() As IFormulas
    Set Self = Me
End Property

'@label:Dictionary
'@sub-title Retrieve the underlying dictionary used during parsing
'@return ILLdictionary previously supplied during Create/Initialise.
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

'@label:Dictionary
'@sub-title Assign a new dictionary and reset dependent caches
'@param dict ILLdictionary providing variables metadata.
Public Property Set Dictionary(ByVal dict As ILLdictionary)
    If dict Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Dictionary is not found"
    Set this.dict = dict
    ResetDependencies
End Property

'@label:Data
'@sub-title Retrieve the formula configuration catalog
'@return IFormulaData describing available Excel tokens.
Public Property Get Data() As IFormulaData
    Set Data = this.data
End Property

'@label:Data
'@sub-title Update the formula configuration catalog and invalidate caches
'@param formData IFormulaData with allowed operators/functions.
Public Property Set Data(ByVal formData As IFormulaData)
    If formData Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Formula Data not found"
    Set this.data = formData
    InvalidateCaches
End Property

'@label:SetupFormula
'@sub-title Retrieve the original setup pseudo-formula
'@return String raw expression.
Public Property Get SetupFormula() As String
    SetupFormula = this.setupForm
End Property

'@label:SetupFormula
'@sub-title Store the setup pseudo-formula and reset computed caches
'@param setupForm String pseudo-formula captured from setup.
Public Property Let SetupFormula(ByVal setupForm As String)
    this.setupForm = setupForm
    InvalidateCaches
End Property

'@label:HasSetupVariables
'@sub-title Determine if the parsed formula contains literal values
'@return Boolean True when numeric/text literals were encountered.
Public Property Get HasSetupVariables() As Boolean
    HasSetupVariables = this.containsSetupVariables
End Property

'@label:IsGrouped
'@sub-title Indicate whether the parsed formula targets grouped evaluation
'@return String "Yes" when grouped logic is applied, "No" otherwise.
Public Property Get IsGrouped() As String
    Call EnsureEvaluation("simple")

    If this.group.Enabled Then
        IsGrouped = "Yes"
    Else
        IsGrouped = "No"
    End If

End Property

'@label:Reason
'@sub-title Retrieve the last validation message associated with the given context
'@param formulaType Optional String context (analysis/simple).
'@return String explanation of the current validity state.
Public Property Get Reason(Optional ByVal formulaType As String = "analysis") As String
    If EnsureEvaluation(formulaType) Then
        If this.isValid Then
            Reason = FORMULA_DEFAULT_REASON
        Else
            Reason = this.invalidReason
        End If
    Else
        Reason = this.invalidReason
    End If
End Property

'@label:Valid
'@sub-title Check whether the formula is valid for the supplied context
'@param formulaType Optional String context (analysis/simple).
'@return Boolean True when parsing succeeded.
Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    Valid = EnsureEvaluation(formulaType)
End Property

'@label:Varlists
'@sub-title Provide a clone of the variables detected in the expression
'@param formulaType Optional String context (analysis/simple).
'@return BetterArray containing variable names.
Public Property Get Varlists(Optional ByVal formulaType As String = "analysis") As BetterArray
    Dim result As BetterArray

    Set result = New BetterArray
    result.LowerBound = 1

    If EnsureEvaluation(formulaType) Then
        If Not this.variables Is Nothing Then
            Set result = this.variables.Clone()
        End If
    End If

    Set Varlists = result.Clone()
End Property

'@label:ParsedLinelistFormula
'@sub-title Build the Excel expression for linelist worksheets
'@param useTableName Optional Boolean toggling structured references.
'@param tablePrefix Optional String prefixed to structured references.
'@return String Excel formula referencing linelist variables.
Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
    Dim workingTokens As BetterArray
    Dim idx As Long
    Dim current As String
    Dim vars As ILLVariables
    Dim sheets As ILLSheets

    If Not EnsureEvaluation("simple") Then Exit Property

    If this.group.Enabled Then
        ParsedLinelistFormula = BuildGroupedFormula(useTableName, tablePrefix)
        Exit Property
    End If

    Set workingTokens = this.tokens.Clone()
    Set vars = Variables()
    Set sheets = SheetProvider()

    For idx = workingTokens.LowerBound To workingTokens.UpperBound
        current = CStr(workingTokens.Item(idx))
        If vars.Contains(current) Then
            If useTableName Then
                workingTokens.Item(idx) = tablePrefix & vars.Value(colName:="table name", varName:=current) & "[" & current & "]"
            Else
                workingTokens.Item(idx) = sheets.VariableAddress(current)
            End If
        ElseIf StrComp(current, "MEAN", vbTextCompare) = 0 Then
            workingTokens.Item(idx) = "AVERAGE"
        ElseIf StrComp(current, "N", vbTextCompare) = 0 Then
            workingTokens.Item(idx) = "COUNT"
        End If
    Next idx

    ParsedLinelistFormula = workingTokens.ToString(Separator:=vbNullString, _
                                                   OpeningDelimiter:=vbNullString, _
                                                   ClosingDelimiter:=vbNullString, _
                                                   QuoteStrings:=False)
End Property

'@label:ParsedAnalysisFormula
'@sub-title Build the Excel expression for aggregated analysis results
'@param formCond IFormulaCondition providing conditional aggregation logic.
'@param tablePrefix Optional String used to prefix table references.
'@param Connector Optional String used between conditional fragments.
'@return String Excel expression ready for analysis worksheets.
Public Property Get ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
    Dim vars As ILLVariables
    Dim dict As ILLdictionary
    Dim tokensCopy As BetterArray
    Dim idx As Long
    Dim chunk As String
    Dim tableName As String
    Dim parsedCustom As String
    Dim lowerBound As Long
    Dim upperBound As Long

    If Not EnsureEvaluation("analysis") Then Exit Property

    If this.group.Enabled Then
        ParsedAnalysisFormula = BuildGroupedFormula(True, tablePrefix)
        Exit Property
    End If

    If formCond Is Nothing Then Exit Property

    Set tokensCopy = this.tokens.Clone()
    Set vars = Variables()
    Set dict = Dictionary()

    lowerBound = tokensCopy.LowerBound
    upperBound = tokensCopy.UpperBound
    idx = lowerBound

    Do While idx <= upperBound
        chunk = CStr(tokensCopy.Item(idx))

        If vars.Contains(chunk) Then
            tableName = vars.Value(varName:=chunk, colName:="table name")
            If Not formCond.Valid(dict, tableName) Then Exit Property
            tokensCopy.Item(idx) = formCond.ConditionString(varName:=chunk, _
                                                            TableName:=tablePrefix & tableName, _
                                                            Connector:=Connector)
        ElseIf deps.customFormulas.Includes(chunk) Then
            parsedCustom = ParsedCustomFormula(customFunction:=chunk, _
                                               formCond:=formCond, _
                                               TableName:=tableName, _
                                               tablePrefix:=tablePrefix)
            If parsedCustom = vbNullString Then Exit Property
            tokensCopy.Item(idx) = parsedCustom
            RemoveEmptyInvocationTokens tokensCopy, idx
            upperBound = tokensCopy.UpperBound
        End If
        idx = idx + 1
    Loop

    ParsedAnalysisFormula = tokensCopy.ToString(Separator:=vbNullString, _
                                                OpeningDelimiter:=vbNullString, _
                                                ClosingDelimiter:=vbNullString, _
                                                QuoteStrings:=False)
End Property

'@section PrivateHelpers
'===============================================================================
'@details Contains helpers that prepare dependencies, tokenise the expression, and materialise parsed results.

'@label:Initialise
'@sub-title Store incoming dependencies and prepare caches
'@param dict ILLdictionary backing metadata lookups.
'@param formData IFormulaData defining allowed tokens/operators.
'@param setupForm String pseudo-formula to parse.
Public Sub Initialise(ByVal dict As ILLdictionary, _
                       ByVal formData As IFormulaData, _
                       ByVal setupForm As String)
    Set this.dict = dict
    Set this.data = formData
    this.setupForm = setupForm
    InvalidateCaches
    ResetDependencies
    RegisterDefaultCustomFormulas
End Sub

'@label:ResetDependencies
'@sub-title Clear lazily instantiated dependencies and invalidate token caches
Private Sub ResetDependencies()
    Set deps.variables = Nothing
    Set deps.sheets = Nothing
    If deps.customFormulas Is Nothing Then
        Set deps.customFormulas = New BetterArray
        deps.customFormulas.LowerBound = 0
    End If
    InvalidateCaches
End Sub

'@label:RegisterDefaultCustomFormulas
'@sub-title Seed built-in custom formulas recognised by analysis parsing
Private Sub RegisterDefaultCustomFormulas()
    Set deps.customFormulas = New BetterArray
    deps.customFormulas.LowerBound = 0
    RegisterCustomFormula "MEAN"
    RegisterCustomFormula "N"
    RegisterCustomFormula "N()"
End Sub

'@label:RegisterCustomFormula
'@sub-title Add a custom formula identifier if not already registered
'@param customFormula String token representing a custom formula.
Private Sub RegisterCustomFormula(ByVal customFormula As String)
    If LenB(customFormula) = 0 Then Exit Sub
    If deps.customFormulas Is Nothing Then
        Set deps.customFormulas = New BetterArray
        deps.customFormulas.LowerBound = 0
    End If
    If Not deps.customFormulas.Includes(customFormula) Then deps.customFormulas.Push customFormula
End Sub

'@label:InvalidateCaches
'@sub-title Reset cached parsing results and computed flags
Private Sub InvalidateCaches()
    Set this.tokens = Nothing
    Set this.variables = Nothing
    this.containsSetupVariables = False
    this.isValid = False
    this.invalidReason = vbNullString
    this.validityComputed = False
    this.lastFormulaType = vbNullString
    ResetGroupContext
End Sub

'@label:ResetGroupContext
'@sub-title Clear any grouped-formula metadata captured during previous evaluations
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

'@label:Variables
'@sub-title Lazily instantiate the LLVariables helper from the dictionary
'@return ILLVariables bound to the current dictionary.
Private Function Variables() As ILLVariables
    If deps.variables Is Nothing Then
        Set deps.variables = LLVariables.Create(Dictionary())
    End If
    Set Variables = deps.variables
End Function

'@label:SheetProvider
'@sub-title Lazily instantiate the LLSheets helper for address resolution
'@return ILLSheets bound to the current dictionary.
Private Function SheetProvider() As ILLSheets
    If deps.sheets Is Nothing Then
        Set deps.sheets = LLSheets.Create(Dictionary())
    End If
    Set SheetProvider = deps.sheets
End Function

'@label:EnsureEvaluation
'@sub-title Guarantee that a parsing/validation run exists for the requested context
'@param formulaType String context (analysis/simple).
'@return Boolean True when the cached evaluation is valid.
Private Function EnsureEvaluation(ByVal formulaType As String) As Boolean
    If Not this.validityComputed Or StrComp(this.lastFormulaType, formulaType, vbTextCompare) <> 0 Then
        EvaluateFormula formulaType
    End If
    EnsureEvaluation = this.isValid
End Function

'@label:EvaluateFormula
'@sub-title Parse the setup expression, populate tokens, and capture validation results
'@param formulaType String context (analysis/simple).
Private Sub EvaluateFormula(ByVal formulaType As String)
    Dim converted As String
    Dim vars As ILLVariables
    Dim dataSource As IFormulaData
    Dim tokens As BetterArray
    Dim variableTokens As BetterArray
    Dim formulaHasSetupVariables As Boolean
    Dim reasonText As String

    InvalidateCaches
    this.lastFormulaType = formulaType


    converted = ConvertedSetupFormula()
    converted = CleanString(converted)

    Set vars = Variables()
    Set dataSource = Data()

    If LenB(converted) = 0 Then
        reasonText = FORMULA_EMPTY_MESSAGE
        SetInvalid reasonText
        Exit Sub
    End If

    If vars.Contains(converted) Then
        If StrComp(formulaType, "analysis", vbTextCompare) = 0 Then
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

    If StrComp(formulaType, "analysis", vbTextCompare) = 0 And deps.customFormulas.Includes(converted) Then
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

    '@Ignore AssignmentNotUsed
    Set tokens = New BetterArray
    tokens.LowerBound = 1
    '@Ignore AssignmentNotUsed
    Set variableTokens = New BetterArray
    variableTokens.LowerBound = 1


    If Not TokeniseFormula(converted, formulaType, tokens, variableTokens, formulaHasSetupVariables, reasonText) Then
        SetInvalid reasonText
        Exit Sub
    End If

    If Not GroupContexExtractionSucceed(tokens, variableTokens, vars, dataSource) Then Exit Sub

    '@Ignore UnassignedVariableUsage
    this.containsSetupVariables = formulaHasSetupVariables
    Set this.tokens = tokens
    Set this.variables = variableTokens
    this.isValid = True
    this.validityComputed = True
End Sub

'@section GroupFormula
'===============================================================================
'@details Detects grouped formulas, validates arguments, and prepares metadata
'@details for downstream builders.

'@label:HasGroupContext
'@sub-title Analyse the tokenised expression and record grouped metadata when relevant
'@param tokens BetterArray containing the tokenised formula.
'@param variableTokens BetterArray listing variable names in encounter order.
'@param vars ILLVariables helper exposing dictionary metadata.
'@param dataSource IFormulaData configuration describing grouped functions.
'@return Boolean True when grouped context is captured or not required.
Private Function GroupContexExtractionSucceed(ByVal tokens As BetterArray, _
                                       ByVal variableTokens As BetterArray, _
                                       ByVal vars As ILLVariables, _
                                       ByVal dataSource As IFormulaData) As Boolean
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

'@label:RootFunctionName
'@sub-title Retrieve the entry token corresponding to the root function
'@param tokens BetterArray containing the tokenised expression.
'@return String root token or empty string when unavailable.
Private Function RootFunctionName(ByVal tokens As BetterArray) As String
    If tokens Is Nothing Then Exit Function
    If tokens.Length = 0 Then Exit Function
    RootFunctionName = CStr(tokens.Item(tokens.LowerBound))
End Function

'@label:IsGenericGroupFunction
'@sub-title Determine whether a function uses the GROUP prefix for grouped evaluation
'@param functionName String token under inspection.
'@return Boolean True when the token starts with "GROUP".
Private Function IsGenericGroupFunction(ByVal functionName As String) As Boolean
    If LenB(functionName) = 0 Then Exit Function
    IsGenericGroupFunction = (InStr(1, functionName, "GROUP", vbTextCompare) = 1)
End Function

'@label:ExtractGenericAggregator
'@sub-title Parse the aggregator portion from a GROUP-prefixed function
'@param functionName String grouped token (e.g. GROUP_SUM or GROUPS_SUM)
'@return String aggregator token to apply when building the Excel formula.
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

'@label:BuildGroupedFormula
'@sub-title Construct the grouped Excel expression using captured metadata
'@param useTableName Boolean True to emit structured references.
'@param tablePrefix String prefix applied to structured references when requested.
'@return String grouped Excel formula ready for consumption.
Private Function BuildGroupedFormula(ByVal useTableName As Boolean, _
                                     ByVal tablePrefix As String) As String
    Dim sheets As ILLSheets
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
                BuildGroupedFormula = "COUNTIFS(" & criteriaRange & ", " & conditionValue & ", " & resultRange & ", " & Quoted("<>") & ")"
            Case Else
                BuildGroupedFormula = aggregator & "(" & resultRange & ", " & criteriaRange & ", " & conditionValue & ")"
        End Select
    Else
        BuildGroupedFormula = aggregator & "(IF(" & criteriaRange & "=" & conditionValue & ", " & resultRange & "))"
    End If
End Function

'@label:GroupedDataRange
'@sub-title Resolve the appropriate range reference for a grouped variable
'@param variableName String variable identifier.
'@param useTableName Boolean True to emit structured references.
'@param tablePrefix String prefix applied to structured references.
'@param sheets ILLSheets provider for address resolution.
'@param sheetName String sheet hosting the formula; used for relative addressing.
'@return String Excel reference targeting the grouped variable.
Private Function GroupedDataRange(ByVal variableName As String, _
                                  ByVal useTableName As Boolean, _
                                  ByVal tablePrefix As String, _
                                  ByVal sheets As ILLSheets) As String
    Dim tableName As String

    tableName = this.group.TableName

    If useTableName Then
        GroupedDataRange = tablePrefix & tableName & "[" & variableName & "]"
    Else
        GroupedDataRange = sheets.VariableAddress(variableName)
    End If
End Function

'@label:TokeniseFormula
'@sub-title Break the converted formula into tokens while validating parentheses and literals
'@param formula String cleaned setup expression.
'@param formulaType String context (analysis/simple).
'@param tokens BetterArray recipient for all tokens.
'@param variableTokens BetterArray recipient for variable-only tokens.
'@param HasSetupVariables Boolean flag capturing literal presence.
'@param failureReason String explanation when tokenisation fails.
'@return Boolean True when parsing succeeds.
Private Function TokeniseFormula(ByVal formula As String, _
                                 ByVal formulaType As String, _
                                 ByVal tokens As BetterArray, _
                                 ByVal variableTokens As BetterArray, _
                                 ByRef HasSetupVariables As Boolean, _
                                 ByRef failureReason As String) As Boolean
    Dim idx As Long
    Dim previousBreak As Long
    Dim currentChar As String
    Dim chunk As String
    Dim parenthesisLevel As Long
    Dim quoted As Boolean
    Dim dataSource As IFormulaData
    Dim vars As ILLVariables

    previousBreak = 1
    Set dataSource = Data()
    Set vars = Variables()

    For idx = 1 To Len(formula)
        currentChar = Mid$(formula, idx, 1)

        If currentChar = Chr$(34) Then
            quoted = Not quoted
        ElseIf Not quoted Then
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

        If Not quoted And dataSource.SpecialCharacterIncludes(currentChar) Then
            chunk = CleanString(Mid$(formula, previousBreak, idx - previousBreak))
            If Not AppendToken(chunk, formulaType, tokens, variableTokens, vars, dataSource, HasSetupVariables, failureReason) Then Exit Function
            tokens.Push currentChar
            previousBreak = idx + 1
        End If
    Next idx

    chunk = CleanString(Mid$(formula, previousBreak, Len(formula) - previousBreak + 1))
    If Not AppendToken(chunk, formulaType, tokens, variableTokens, vars, dataSource, HasSetupVariables, failureReason) Then Exit Function

    If parenthesisLevel <> 0 Then
        failureReason = FORMULA_PAREN_MISMATCH
        Exit Function
    End If

    TokeniseFormula = True
End Function

'@label:AppendToken
'@sub-title Validate and append the provided chunk to the tokens collection
'@param chunk String candidate token.
'@param formulaType String context (analysis/simple).
'@param tokens BetterArray recipient for all tokens.
'@param variableTokens BetterArray recipient for variable-only tokens.
'@param vars ILLVariables metadata provider.
'@param dataSource IFormulaData formula metadata provider.
'@param HasSetupVariables Boolean flag capturing literal presence.
'@param failureReason String explanation when parsing fails.
'@return Boolean True when the chunk is accepted.
Private Function AppendToken(ByVal chunk As String, _
                             ByRef formulaType As String, _
                             ByRef tokens As BetterArray, _
                             ByRef variableTokens As BetterArray, _
                             ByRef vars As ILLVariables, _
                             ByRef dataSource As IFormulaData, _
                             ByRef HasSetupVariables As Boolean, _
                             ByRef failureReason As String) As Boolean
    Dim upperChunk As String
    Dim functionName As String

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

    If StrComp(formulaType, "analysis", vbTextCompare) = 0 And deps.customFormulas.Includes(chunk) Then
        tokens.Push chunk
        AppendToken = True
        Exit Function
    End If

    If IsNumeric(chunk) Then
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

'@label:IsQuotedString
'@sub-title Determine whether the supplied value is wrapped in quotes
'@param value String candidate literal.
'@return Boolean True when the string starts and ends with quotes.
Private Function IsQuotedString(ByVal value As String) As Boolean
    If Len(value) < 2 Then Exit Function
    IsQuotedString = (Left$(value, 1) = Chr$(34) And Right$(value, 1) = Chr$(34))
End Function

'@label:CleanString
'@sub-title Normalise the incoming token by removing control characters and duplicate spaces
'@param value String to clean.
'@return String trimmed and cleaned representation.
Private Function CleanString(ByVal value As String) As String
    Dim cleaned As String
    cleaned = Application.WorksheetFunction.Substitute(value, Chr$(160), " ")
    cleaned = Application.WorksheetFunction.Clean(cleaned)
    CleanString = Application.WorksheetFunction.Trim(cleaned)
End Function

'@label:ParsedCustomFormula
'@sub-title Delegate parsing to custom functions when present in analysis context
'@param customFunction String identifier of the custom formula.
'@param formCond IFormulaCondition providing condition logic.
'@param TableName Optional String table name resolved for variables.
'@param tablePrefix Optional String prefix for structured references.
'@return String Excel fragment produced by the custom handler.
Private Function ParsedCustomFormula(ByVal customFunction As String, _
                                     ByVal formCond As IFormulaCondition, _
                                     Optional ByVal TableName As String = vbNullString, _
                                     Optional ByVal tablePrefix As String = vbNullString) As String
    Dim cond As BetterArray
    Dim condVars As BetterArray
    Dim parsedFormula As String
    Dim customTableName As String
    Dim dictObj As ILLdictionary
    Dim idx As Long

    Set dictObj = Dictionary()
    customTableName = TableName
    If customTableName = vbNullString Then customTableName = formCond.VariablesTable(dictObj)
    customTableName = tablePrefix & customTableName

    Select Case UCase$(customFunction)
        Case "N", "N()"
            Set cond = formCond.Conditions()
            Set condVars = formCond.Variables()
            parsedFormula = customTableName & "[" & condVars.Item(condVars.LowerBound) & "]" & ", " & ClearCountIf(cond.Item(condVars.LowerBound))
            For idx = condVars.LowerBound + 1 To condVars.UpperBound
                parsedFormula = parsedFormula & "," & customTableName & "[" & condVars.Item(idx) & "]" & "," & ClearCountIf(cond.Item(idx))
            Next idx
            parsedFormula = "COUNTIFS(" & parsedFormula & ")"
        Case "MEAN"
            parsedFormula = "AVERAGE"
        Case Else
            parsedFormula = customFunction
    End Select

    ParsedCustomFormula = parsedFormula
End Function

'@label:RemoveEmptyInvocationTokens
'@sub-title Strip trailing empty parentheses after replacing custom formulas
'@param tokens BetterArray tokens list to mutate.
'@param functionIndex Long index of the custom function token.
Private Sub RemoveEmptyInvocationTokens(ByRef tokens As BetterArray, _
                                        ByVal functionIndex As Long)
    Dim openIndex As Long
    Dim closeIndex As Long
    Dim upperBound As Long
    Dim openToken As String
    Dim closeToken As String

    If tokens Is Nothing Then Exit Sub

    upperBound = tokens.UpperBound
    openIndex = functionIndex + 1
    closeIndex = functionIndex + 2

    If openIndex > upperBound Then Exit Sub
    openToken = CStr(tokens.Item(openIndex))
    If StrComp(openToken, "(", vbBinaryCompare) <> 0 Then Exit Sub

    If closeIndex > tokens.UpperBound Then Exit Sub
    closeToken = CStr(tokens.Item(closeIndex))
    If StrComp(closeToken, ")", vbBinaryCompare) <> 0 Then Exit Sub

    Call tokens.Remove(closeIndex)
    Call tokens.Remove(openIndex)
End Sub

'@label:ClearCountIf
'@sub-title Remove equality fragments from COUNTIF-style conditions
'@param value String predicate fragment.
'@return String sanitized predicate suitable for COUNTIFS.
Private Function ClearCountIf(ByVal value As String) As String
    Dim cleared As String
    cleared = Replace(value, Quoted("=") & "&", vbNullString)
    cleared = Replace(cleared, "=", vbNullString)
    cleared = Replace(cleared, Chr$(34) & Chr$(34) & Chr$(34) & Chr$(34), Chr$(34) & Chr$(34))
    cleared = Replace(cleared, "<>", "######")
    cleared = Replace(cleared, "<", Quoted("<=") & " & ")
    cleared = Replace(cleared, ">", Quoted(">=") & " & ")
    cleared = Replace(cleared, "######", "<>")
    ClearCountIf = cleared
End Function

'@label:Quoted
'@sub-title Wrap a value in Excel double quotes
'@param value String to wrap.
'@return String quoted representation.
Private Function Quoted(ByVal value As String) As String
    Quoted = Chr$(34) & value & Chr$(34)
End Function

'@label:ConvertedSetupFormula
'@sub-title Convert higher-level setup directives into plain Excel expressions
'@return String converted formula ready for tokenising.
Private Function ConvertedSetupFormula() As String
    Dim result As String
    Dim caseWhenObj As ICaseWhen
    Dim choiceObj As IChoiceFormula
    Dim valueObj As IValueOfFormula

    result = SetupFormula()

    If InStr(1, result, "CASE_WHEN", vbTextCompare) = 1 Then
        Set caseWhenObj = CaseWhen.Create(result)
        result = caseWhenObj.parsedFormula()
    ElseIf InStr(1, result, "CHOICE_FORMULA", vbTextCompare) = 1 Then
        Set choiceObj = ChoiceFormula.Create(result)
        result = choiceObj.parsedFormula()
    ElseIf InStr(1, result, "VALUE_OF", vbTextCompare) = 1 Then
        Set valueObj = ValueOfFormula.Create(result, Dictionary(), Variables(), SheetProvider())
        result = valueObj.ConvertedFormula
    End If

    ConvertedSetupFormula = result
End Function

'@label:SetInvalid
'@sub-title Record an invalidation reason and log the associated checking entry
'@param message String explaining the validation failure.
'@param scope Optional Byte severity level persisted in Checking.
Private Sub SetInvalid(ByVal message As String, Optional ByVal scope As Byte = checkingError)
    this.invalidReason = message
    this.isValid = False
    this.validityComputed = True
    LogCheck message, scope
End Sub

'@section Checkings
'===============================================================================
'@details Aggregates informational/error traces emitted during parsing.

'@label:LogCheck
'@sub-title Record a diagnostic entry within the internal checking store
'@param label String message to log.
'@param scope Optional Byte severity indicator.
Private Sub LogCheck(ByVal label As String, _
                     Optional ByVal scope As Byte = checkingNote)
    
    If internalChecks Is Nothing Then 
        Set internalChecks = Checking.Create("Formulas")
    End If
    
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    hasChecks = True
End Sub

'@label:HasCheckings
'@sub-title Indicate whether any diagnostic entries were recorded
'@return Boolean True when checkings exist.
Private Property Get HasCheckings() As Boolean
    HasCheckings = hasChecks
End Property

'@label:CheckingValues
'@sub-title Expose the collected checking entries if available
'@return IChecking containing logged diagnostics.
Private Property Get CheckingValues() As IChecking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property

'@section ErrorHandling
'===============================================================================
'@details Centralises the error raising pattern for consistent ProjectError usage.

'@label:ThrowError
'@sub-title Raise a VBA error with the class signature
'@param errNumb Long ProjectError code.
'@param errorMessage String descriptive message.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error: " & errorMessage
End Sub

'@section Interface
'===============================================================================

Private Property Get IFormulas_Varlists(Optional ByVal formulaType As String = "analysis") As BetterArray
    Set IFormulas_Varlists = Varlists(formulaType)
End Property

Private Property Get IFormulas_Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    IFormulas_Valid = Valid(formulaType)
End Property

Private Property Get IFormulas_FormulaTableList() As BetterArray
    If EnsureEvaluation("analysis") Then
        Set IFormulas_FormulaTableList = this.tokens.Clone()
    Else
        Set IFormulas_FormulaTableList = New BetterArray
    End If
End Property

Private Property Get IFormulas_ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                                    Optional ByVal tablePrefix As String = vbNullString) As String
    IFormulas_ParsedLinelistFormula = ParsedLinelistFormula(useTableName, tablePrefix)
End Property

Private Property Get IFormulas_ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, _
                                                    Optional ByVal tablePrefix As String = vbNullString, _
                                                    Optional ByVal Connector As String = "*") As String
    IFormulas_ParsedAnalysisFormula = ParsedAnalysisFormula(formCond, tablePrefix, Connector)
End Property

Private Property Get IFormulas_HasSetupVariables() As Boolean
    IFormulas_HasSetupVariables = HasSetupVariables()
End Property

Private Property Get IFormulas_IsGrouped() As String
    IFormulas_IsGrouped = IsGrouped()
End Property

Private Property Get IFormulas_Reason(Optional ByVal formulaType As String = "analysis") As String
    IFormulas_Reason = Reason(formulaType)
End Property

Private Property Get IFormulas_CheckingValues() As IChecking
    Set IFormulas_CheckingValues = CheckingValues
End Property

Private Property Get IFormulas_HasChecking() As Boolean
    IFormulas_HasChecking = HasCheckings
End Property

