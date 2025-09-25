VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Convert formulas from setup to excel"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Convert formulas from setup to Excel")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, ReadOnlyPropertyAssignment, UseMeaningfulName, HungarianNotation
'@depends BetterArray, Checking, LLdictionary, LLVariables, LLSheets, CaseWhen, ChoiceFormula

Option Explicit
Implements IFormulas

'@section State
'===============================================================================

Private Type TFormulas
    dict As ILLdictionary
    data As IFormulaData
    setupForm As String
    tokens As BetterArray
    variables As BetterArray
    hasLiterals As Boolean
    isValid As Boolean
    invalidReason As String
    validityComputed As Boolean
    lastFormulaType As String
End Type

Private Type TDependencies
    variables As ILLVariables
    sheets As ILLSheets
    customFormulas As BetterArray
End Type

Private Const CLASSNAME As String = "Formulas"
Private Const FORMULA_DEFAULT_REASON As String = "The formula seems correct"
Private Const FORMULA_EMPTY_MESSAGE As String = "The formula is empty. No formula were found"
Private Const FORMULA_ANALYSIS_SINGLE_VAR As String = "Analysis formula can not consist of only one variable, you should use aggregation function"
Private Const FORMULA_UNKNOWN_TOKEN As String = "Unknown token '%1' encountered while parsing"
Private Const FORMULA_PAREN_MISMATCH As String = "The formula contains unmatched parentheses"
Private Const FORMULA_NEGATIVE_PAREN As String = "Closing parenthesis detected before opening one"

Private this As TFormulas
Private deps As TDependencies
Private internalChecks As IChecking
Private checkCounter As Long
Private hasChecks As Boolean

'@section Factory
'===============================================================================

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
    If dict Is Nothing Then ThrowError(ProjectError.ObjectNotInitialized, "Dictionary is not found")
    If formData Is Nothing Then ThrowError(ProjectError.ObjectNotInitialized, "Formula Data not found")

    With New Formulas
        .Initialise dict, formData, setupForm
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================

Public Property Get Self() As IFormulas
    Set Self = Me
End Property

Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.dict
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
    If dict Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Dictionary is not found"
    Set this.dict = dict
    ResetDependencies
End Property

Public Property Get Data() As IFormulaData
    Set Data = this.data
End Property

Public Property Set Data(ByVal formData As IFormulaData)
    If formData Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Formula Data not found"
    Set this.data = formData
    InvalidateCaches
End Property

Public Property Get SetupFormula() As String
    SetupFormula = this.setupForm
End Property

Public Property Let SetupFormula(ByVal setupForm As String)
    this.setupForm = setupForm
    InvalidateCaches
End Property

Public Property Get HasLiterals() As Boolean
    EnsureEvaluation "analysis"
    HasLiterals = this.hasLiterals
End Property

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

Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean
    Valid = EnsureEvaluation(formulaType)
End Property

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

Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
    Dim workingTokens As BetterArray
    Dim idx As Long
    Dim current As String
    Dim vars As ILLVariables
    Dim sheets As ILLSheets
    Dim sheetName As String

    If Not EnsureEvaluation("simple") Then Exit Property

    Set workingTokens = this.tokens.Clone()
    Set vars = Variables()
    Set sheets = SheetProvider()

    sheetName = FormulaSheet()

    For idx = workingTokens.LowerBound To workingTokens.UpperBound
        current = CStr(workingTokens.Item(idx))
        If vars.Contains(current) Then
            If useTableName Then
                workingTokens.Item(idx) = tablePrefix & vars.Value(colName:="table name", varName:=current) & "[" & current & "]"
            Else
                workingTokens.Item(idx) = sheets.VariableAddress(current, sheetName)
            End If
        ElseIf StrComp(current, "MEAN", vbTextCompare) = 0 Then
            workingTokens.Item(idx) = "AVERAGE"
        End If
    Next idx

    ParsedLinelistFormula = workingTokens.ToString(Separator:=vbNullString, _
                                                   OpeningDelimiter:=vbNullString, _
                                                   ClosingDelimiter:=vbNullString, _
                                                   QuoteStrings:=False)
End Property

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

    If formCond Is Nothing Then Exit Property
    If Not EnsureEvaluation("analysis") Then Exit Property

    Set tokensCopy = this.tokens.Clone()
    Set vars = Variables()
    Set dict = Dictionary()

    For idx = tokensCopy.LowerBound To tokensCopy.UpperBound
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
        End If
    Next idx

    ParsedAnalysisFormula = tokensCopy.ToString(Separator:=vbNullString, _
                                                OpeningDelimiter:=vbNullString, _
                                                ClosingDelimiter:=vbNullString, _
                                                QuoteStrings:=False)
End Property

'@section PrivateHelpers
'===============================================================================

Private Sub Initialise(ByVal dict As ILLdictionary, _
                       ByVal formData As IFormulaData, _
                       ByVal setupForm As String)
    Set this.dict = dict
    Set this.data = formData
    this.setupForm = setupForm
    InvalidateCaches
    ResetDependencies
    RegisterDefaultCustomFormulas
End Sub

Private Sub ResetDependencies()
    Set deps.variables = Nothing
    Set deps.sheets = Nothing
    If deps.customFormulas Is Nothing Then
        Set deps.customFormulas = New BetterArray
        deps.customFormulas.LowerBound = 0
    End If
    InvalidateCaches
End Sub

Private Sub RegisterDefaultCustomFormulas()
    Set deps.customFormulas = New BetterArray
    deps.customFormulas.LowerBound = 0
    RegisterCustomFormula "MEAN"
    RegisterCustomFormula "N"
    RegisterCustomFormula "N()"
End Sub

Private Sub RegisterCustomFormula(ByVal customFormula As String)
    If LenB(customFormula) = 0 Then Exit Sub
    If deps.customFormulas Is Nothing Then
        Set deps.customFormulas = New BetterArray
        deps.customFormulas.LowerBound = 0
    End If
    If Not deps.customFormulas.Includes(customFormula) Then deps.customFormulas.Push customFormula
End Sub

Private Sub InvalidateCaches()
    Set this.tokens = Nothing
    Set this.variables = Nothing
    this.hasLiterals = False
    this.isValid = False
    this.invalidReason = vbNullString
    this.validityComputed = False
    this.lastFormulaType = vbNullString
End Sub

Private Function Variables() As ILLVariables
    If deps.variables Is Nothing Then
        Set deps.variables = LLVariables.Create(Dictionary())
    End If
    Set Variables = deps.variables
End Function

Private Function SheetProvider() As ILLSheets
    If deps.sheets Is Nothing Then
        Set deps.sheets = LLSheets.Create(Dictionary())
    End If
    Set SheetProvider = deps.sheets
End Function

Private Function EnsureEvaluation(ByVal formulaType As String) As Boolean
    If Not this.validityComputed Or StrComp(this.lastFormulaType, formulaType, vbTextCompare) <> 0 Then
        EvaluateFormula formulaType
    End If
    EnsureEvaluation = this.isValid
End Function

Private Sub EvaluateFormula(ByVal formulaType As String)
    Dim converted As String
    Dim vars As ILLVariables
    Dim dataSource As IFormulaData
    Dim tokens As BetterArray
    Dim variableTokens As BetterArray
    Dim hasLiterals As Boolean
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

            this.hasLiterals = True
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
        this.hasLiterals = True
        this.isValid = True
        this.validityComputed = True
        Exit Sub
    End If

    Set tokens = New BetterArray
    tokens.LowerBound = 1
    Set variableTokens = New BetterArray
    variableTokens.LowerBound = 1

    If Not TokeniseFormula(converted, formulaType, tokens, variableTokens, hasLiterals, reasonText) Then
        SetInvalid reasonText
        Exit Sub
    End If

    this.hasLiterals = hasLiterals
    Set this.tokens = tokens
    Set this.variables = variableTokens
    this.isValid = True
    this.validityComputed = True
End Sub

Private Function TokeniseFormula(ByVal formula As String, _
                                 ByVal formulaType As String, _
                                 ByVal tokens As BetterArray, _
                                 ByVal variableTokens As BetterArray, _
                                 ByRef hasLiterals As Boolean, _
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
            If Not AppendToken(chunk, formulaType, tokens, variableTokens, vars, dataSource, hasLiterals, failureReason) Then Exit Function
            tokens.Push currentChar
            previousBreak = idx + 1
        End If
    Next idx

    chunk = CleanString(Mid$(formula, previousBreak, Len(formula) - previousBreak + 1))
    If Not AppendToken(chunk, formulaType, tokens, variableTokens, vars, dataSource, hasLiterals, failureReason) Then Exit Function

    If parenthesisLevel <> 0 Then
        failureReason = FORMULA_PAREN_MISMATCH
        Exit Function
    End If

    TokeniseFormula = True
End Function

Private Function AppendToken(ByVal chunk As String, _
                             ByVal formulaType As String, _
                             ByVal tokens As BetterArray, _
                             ByVal variableTokens As BetterArray, _
                             ByVal vars As ILLVariables, _
                             ByVal dataSource As IFormulaData, _
                             ByRef hasLiterals As Boolean, _
                             ByRef failureReason As String) As Boolean
    Dim upperChunk As String

    If LenB(chunk) = 0 Then
        AppendToken = True
        Exit Function
    End If

    upperChunk = UCase$(chunk)

    If vars.Contains(chunk) Then
        variableTokens.Push chunk
        tokens.Push chunk
        hasLiterals = True
        AppendToken = True
        Exit Function
    End If

    If StrComp(upperChunk, "TRUE", vbTextCompare) = 0 Or _
       StrComp(upperChunk, "FALSE", vbTextCompare) = 0 Then
        tokens.Push upperChunk
        hasLiterals = True
        AppendToken = True
        Exit Function
    End If

    If dataSource.ExcelFormulasIncludes(upperChunk) Then
        tokens.Push upperChunk
        hasLiterals = True
        AppendToken = True
        Exit Function
    End If

    If StrComp(formulaType, "analysis", vbTextCompare) = 0 And deps.customFormulas.Includes(chunk) Then
        tokens.Push chunk
        hasLiterals = True
        AppendToken = True
        Exit Function
    End If

    If IsNumeric(chunk) Then
        tokens.Push chunk
        hasLiterals = True
        AppendToken = True
        Exit Function
    End If

    If IsQuotedString(chunk) Then
        tokens.Push chunk
        hasLiterals = True
        AppendToken = True
        Exit Function
    End If

    failureReason = Replace(FORMULA_UNKNOWN_TOKEN, "%1", chunk)
    AppendToken = False
End Function

Private Function IsQuotedString(ByVal value As String) As Boolean
    If Len(value) < 2 Then Exit Function
    IsQuotedString = (Left$(value, 1) = Chr$(34) And Right$(value, 1) = Chr$(34))
End Function

Private Function CleanString(ByVal value As String) As String
    Dim cleaned As String
    cleaned = Application.WorksheetFunction.Substitute(value, Chr$(160), " ")
    cleaned = Application.WorksheetFunction.Clean(cleaned)
    CleanString = Application.WorksheetFunction.Trim(cleaned)
End Function

Private Function FormulaSheet() As String
    Dim dictObj As ILLdictionary
    Dim sheetList As BetterArray

    Set dictObj = Dictionary()
    FormulaSheet = vbNullString

    Set sheetList = dictObj.Data.FilterData("control details", SetupFormula(), "sheet name")
    If sheetList.Length = 1 Then FormulaSheet = CStr(sheetList.Item(sheetList.LowerBound))
End Function

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

Private Function Quoted(ByVal value As String) As String
    Quoted = Chr$(34) & value & Chr$(34)
End Function

Private Function ConvertedSetupFormula() As String
    Dim result As String
    Dim caseWhenObj As ICaseWhen
    Dim choiceObj As IChoiceFormula

    result = SetupFormula()

    If InStr(1, result, "CASE_WHEN", vbTextCompare) = 1 Then
        Set caseWhenObj = CaseWhen.Create(result)
        result = caseWhenObj.parsedFormula()
    ElseIf InStr(1, result, "CHOICE_FORMULA", vbTextCompare) = 1 Then
        Set choiceObj = ChoiceFormula.Create(result)
        result = choiceObj.parsedFormula()
    End If

    ConvertedSetupFormula = result
End Function

Private Sub SetInvalid(ByVal message As String, Optional ByVal scope As Byte = checkingError)
    this.invalidReason = message
    this.isValid = False
    this.validityComputed = True
    LogCheck message, scope
End Sub

'@section Checkings
'===============================================================================

Private Sub LogCheck(ByVal label As String, _
                     Optional ByVal scope As Byte = checkingNote)
    Dim checkName As String

    checkName = "Formulas " & SetupFormula() & " trace"

    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)
    internalChecks.Add CStr(checkCounter), label, scope
    checkCounter = checkCounter + 1
    hasChecks = True
End Sub

Private Property Get HasCheckings() As Boolean
    HasCheckings = hasChecks
End Property

Private Property Get CheckingValues() As IChecking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = internalChecks
End Property

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumb As ProjectError, ByVal errorMessage As String)
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

Private Property Get IFormulas_HasLiterals() As Boolean
    IFormulas_HasLiterals = HasLiterals()
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

