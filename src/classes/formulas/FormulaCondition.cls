VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Conditions used to create formulas for analysis"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Conditions used to create formulas for analysis")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class FormulaCondition
'@description
'Couples variable names with their condition fragments, validates alignment
'against the dictionary, and produces structured Excel predicates for
'analysis formulas. The class pairs two BetterArray lists (variables and
'conditions), validates that all variables belong to the same table, and
'exposes methods to render IF-wrapped condition strings and joined predicates.
'@depends BetterArray, Checking, LLVariables, LLdictionary
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Tracks the paired variable/condition lists alongside cached diagnostics.

Private Type TFormulaCondition
    condList As BetterArray                     'Conditions for the formulas
    varList As BetterArray                      'Variables for the formulas
    cachedTableName As String                    'Validated table name across variables
    hasChecks As Boolean
    internalChecks As Checking
    checkCounter As Long
    cachedValidDone As Boolean                   'Whether Valid has produced a cached result
    cachedValidResult As Boolean                 'Cached True/False from last Valid call
    cachedValidTable As String                   'The tablename parameter used for cached result
    cachedValidDict As LLdictionary              'The dictionary the cached result belongs to
    varsHelper As LLVariables                    'Helper built once per dictionary
    varsHelperDict As LLdictionary               'The dictionary the helper was built over
    cachedPredicateKey As String                 'tablename and connector the predicate belongs to
    cachedPredicate As String                    'Last predicate built
    cachedPredicateDone As Boolean               'Whether a predicate has been built
    sealed As Boolean
End Type


Private this As TFormulaCondition

Private Const CLASSNAME As String = "FormulaCondition"

'Separates the two parts of the predicate cache key. A table name and a
'connector cannot hold this character.
Private Const KEY_SEPARATOR As String = vbNullChar

'@section Factory
'===============================================================================
'@description Validate paired arrays and expose the ready-to-use condition helper.

'@label:create
'@sub-title Construct a FormulaCondition instance from paired arrays
'@details
'Validates that both arrays are non-Nothing and have equal lengths, then
'stores normalised copies with a lower bound of 1.
'@param formVar BetterArray. The array containing variable names.
'@param formCond BetterArray. The array containing predicate fragments.
'@return FormulaCondition. A ready-to-use condition helper.
'@throws ProjectError.InvalidArgument When formVar or formCond is Nothing or lengths differ.
'@export
Public Function Create(ByVal formVar As BetterArray, ByVal formCond As BetterArray) As FormulaCondition

    If formVar Is Nothing Then ThrowError ProjectError.InvalidArgument, "Variables array cannot be null"
    If formCond Is Nothing Then ThrowError ProjectError.InvalidArgument, "Conditions array cannot be null"

    If formVar.Length <> formCond.Length Then _
        ThrowError ProjectError.InvalidArgument, _
                  "Condition length " & formCond.Length & " different from variable length " & formVar.Length

    Dim instance As FormulaCondition
    Set instance = New FormulaCondition
    Set instance.LstConditions = formCond
    Set instance.LstVars = formVar
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
'@description Provide controlled access to stored collections.

'@label:conditions
'@prop-title Retrieve the stored condition fragments
'@details
'Returns the internal BetterArray holding the condition predicates
'in the same order they were supplied at creation.
'@return BetterArray. The condition fragments preserving the original order.
'@export
Public Property Get LstConditions() As BetterArray
    Set LstConditions = this.condList
End Property

'@label:conditions-set
'@prop-title Store the condition fragments and reset caches
'@details
'Copies the incoming array into a new BetterArray normalised to a
'lower bound of 1, then resets all cached validation artefacts so
'the next call to Valid triggers a fresh check.
'Both setters copy with LowerBound 1, so the two lists always share one
'lower bound. ConditionPredicate indexes both lists from that bound and it
'is correct because of this.
'@param formCond BetterArray. The array containing predicate fragments.
Public Property Set LstConditions(ByVal formCond As BetterArray)
    GuardNotSealed "LstConditions"
    Set this.condList = New BetterArray
    this.condList.LowerBound = 1
    this.condList.Items = formCond.Items
    ResetCaches
End Property

'@label:variables
'@prop-title Retrieve the stored variables list
'@details
'Returns the internal BetterArray holding the variable names in the
'same order they were supplied at creation.
'@return BetterArray. The variable names preserving the original order.
'@export
Public Property Get LstVars() As BetterArray
    Set LstVars = this.varList
End Property

'@label:variables-set
'@prop-title Store the variables list, normalise bounds, and reset caches
'@details
'Copies the incoming array into a new BetterArray normalised to a
'lower bound of 1, then invalidates all cached data.
'@param formVar BetterArray. The array containing variable identifiers.
Public Property Set LstVars(ByVal formVar As BetterArray)
    GuardNotSealed "LstVars"
    Set this.varList = New BetterArray
    this.varList.LowerBound = 1
    this.varList.Items = formVar.Items
    ResetCaches
End Property

'@section InternalState
'===============================================================================
'@description Maintain cached validation artefacts and diagnostics.

'@label:reset-caches
'@sub-title Clear cached table names, validation results, and diagnostics
'@details
'Resets all memoised state so the next call to Valid, VariablesTable or
'ConditionPredicate starts fresh. Its two callers are the setters above,
'and both are guarded by Seal, so this runs at construction only.
Private Sub ResetCaches()
    this.cachedTableName = vbNullString
    this.cachedValidDone = False
    this.cachedValidResult = False
    this.cachedValidTable = vbNullString
    Set this.cachedValidDict = Nothing
    Set this.varsHelper = Nothing
    Set this.varsHelperDict = Nothing
    this.cachedPredicateKey = vbNullString
    this.cachedPredicate = vbNullString
    this.cachedPredicateDone = False
    ClearCheckings
End Sub

'@label:clear-checkings
'@sub-title Drop the diagnostic entries of the previous validation run
Private Sub ClearCheckings()
    this.hasChecks = False
    this.checkCounter = 0
    Set this.internalChecks = Nothing
End Sub

'@label:log-info
'@sub-title Append a diagnostic entry to the internal checking log
'@details
'Creates the checking instance on first use and appends the message
'using a monotonically increasing counter key. The log is what says WHY
'a formula failed: the analysis path gives up quietly when Valid is False,
'and the reason lives here.
'@param message String. The event description to record.
'@param scope Optional Byte. Severity level from CheckingLogType. Defaults to checkingInfo.
Private Sub LogInfo(ByVal message As String, Optional ByVal scope As Byte = checkingInfo)
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create("FormulaCondition trace")
    this.internalChecks.Add CStr(this.checkCounter), message, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Determine whether diagnostics have been captured
'@details
'Returns True when at least one LogInfo call has been made since the
'last validation run, indicating that diagnostic information is available.
'@return Boolean. True when internal checks exist.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@section MetadataHelpers
'===============================================================================
'@description Provide metadata derived from the last validation run.

'@label:variables-table
'@prop-title Resolve the table name shared by all tracked variables
'@details
'Returns the cached table name when validation has already occurred.
'When the cache is empty, triggers a full validation pass against
'the supplied dictionary before returning the result.
'Formulas.ParsedCustomFormula calls this as a fallback when the table name
'is unknown, so it sits on the hot path for N and N() formulas.
'@param dict LLdictionary. The dictionary used when validation has not occurred yet.
'@return String. Table identifier or empty string when unavailable.
'@export
Public Property Get VariablesTable(ByVal dict As LLdictionary) As String
    If LenB(this.cachedTableName) = 0 Then
        If Not Valid(dict) Then Exit Property
    End If

    VariablesTable = this.cachedTableName
End Property

'@section Validation
'===============================================================================
'@description Confirm that variables/conditions are aligned and belong to the same table.

'@label:valid
'@sub-title Validate variable/condition lists against the dictionary
'@details
'Checks that the lists are non-empty, then verifies each variable exists in
'the dictionary and belongs to the expected table. When a tablename argument
'is supplied it overrides the table resolved from the first variable.
'The result is cached against BOTH the tablename argument and the dictionary
'object. Keying on the tablename alone handed the answer computed for one
'dictionary back to a caller asking about another one.
'@param dict LLdictionary. The metadata provider used for variable lookups.
'@param tablename Optional String. When supplied, forces the expected table name. Defaults to vbNullString.
'@return Boolean. True when validation succeeds.
'@export
Public Function Valid(ByVal dict As LLdictionary, Optional ByVal tablename As String = vbNullString) As Boolean

    Dim vars As LLVariables
    Dim varList As BetterArray
    Dim idx As Long
    Dim lower As Long
    Dim varName As String
    Dim targetTable As String
    Dim currentTable As String

    If this.cachedValidDone Then
        If StrComp(tablename, this.cachedValidTable, vbBinaryCompare) = 0 Then
            If this.cachedValidDict Is dict Then
                Valid = this.cachedValidResult
                Exit Function
            End If
        End If
    End If

    'Fresh validation: drop the previous answer and the previous diagnostics.
    this.cachedTableName = vbNullString
    ClearCheckings

    'Mark cache as populated (default result is False)
    this.cachedValidDone = True
    this.cachedValidTable = tablename
    Set this.cachedValidDict = dict
    this.cachedValidResult = False

    Set varList = this.varList

    'Create raises on a length mismatch and both setters are sealed, so the
    'two lists cannot get out of step after construction. Only the empty
    'case is left to test.
    If varList.Length = 0 Then
        LogInfo "Validation failed: empty variable or condition list", checkingError
        Exit Function
    End If

    On Error GoTo ValidationError
        Set vars = VariablesProvider(dict)
    On Error GoTo 0

    targetTable = tablename
    lower = varList.LowerBound

    For idx = lower To varList.UpperBound

        varName = CStr(varList.Item(idx))

        'Contains is answered from the name block LLVariables holds, so it
        'costs no worksheet read. It stays because a missing variable and a
        'variable with no table need different messages.
        If Not vars.Contains(varName) Then
            LogInfo "Validation failed: variable not found - " & varName, checkingError
            Exit Function
        End If

        currentTable = vars.TableName(varName)

        If idx = lower And LenB(targetTable) = 0 Then
            targetTable = currentTable
            If LenB(targetTable) = 0 Then
                LogInfo "Validation failed: table name not resolved", checkingError
                Exit Function
            End If
        End If

        If StrComp(currentTable, targetTable, vbBinaryCompare) <> 0 Then
            LogInfo "Validation failed: variable belongs to different table - " & varName, checkingError
            Exit Function
        End If
    Next idx

    this.cachedTableName = targetTable
    this.cachedValidResult = True
    Valid = True
    Exit Function

ValidationError:
    'A raise from LLVariables.Create is cached as an invalid formula, which
    'is the intended answer: the dictionary could not be read.
    LogInfo "Validation failed: " & Err.Description, checkingError
    Err.Clear
End Function

'@label:variables-provider
'@sub-title Build the variables helper once per dictionary
'@details
'Valid used to build a fresh LLVariables on every call, and
'Formulas.ParsedAnalysisFormula calls Valid once per variable token of every
'analysis formula. Each helper carries its own block reads over the same
'dictionary worksheet. The helper is held against the dictionary it was
'built over, so a call naming another dictionary builds a new one.
'@param dict LLdictionary. The dictionary to read.
'@return LLVariables. Helper bound to that dictionary.
Private Function VariablesProvider(ByVal dict As LLdictionary) As LLVariables
    If this.varsHelper Is Nothing Or Not (this.varsHelperDict Is dict) Then
        Set this.varsHelper = LLVariables.Create(dict)
        Set this.varsHelperDict = dict
    End If

    Set VariablesProvider = this.varsHelper
End Function

'@section FormulaRendering
'===============================================================================
'@description Produce Excel-ready predicates and wrapper formulas.

'@label:condition-string
'@sub-title Wrap the predicate in an IF expression referencing the supplied variable
'@details
'Joins all variable/condition pairs using ConditionPredicate, then
'wraps the result in an IF function that returns the variable value
'when the predicate evaluates to True. Returns an empty string when
'the predicate is empty.
'The space before the comma is deliberate: CrossTableFormula compares
'generated formulas in its tests.
'@param tablename String. Structured reference table name.
'@param varName String. Variable/column name to return when the condition is met.
'@param Connector Optional String. Connector used between predicates. Defaults to "*".
'@return String. An IF expression or empty string when the predicate is empty.
'@export
Public Function ConditionString(ByVal tablename As String, ByVal varName As String, _
                                Optional ByVal Connector As String = "*") As String

    Dim predicate As String

    predicate = ConditionPredicate(tablename, Connector)
    If LenB(predicate) = 0 Then Exit Function

    ConditionString = "IF(" & predicate & " , " & tablename & "[" & varName & "]" & ")"
End Function

'@label:condition-predicate
'@sub-title Join the variable predicates using the supplied connector
'@details
'Iterates through all variable/condition pairs and concatenates them
'into a single predicate string using the supplied connector (typically
'"*" for AND logic). Each fragment is wrapped in parentheses using
'structured table references, and the condition fragment carries its own
'operator: ">= 5" becomes (f_table[age]>= 5). Formulas.ClearCountIf rewrites
'those fragments for COUNTIFS, so the two have to keep agreeing on the shape.
'The answer is held against the table name and the connector. The lists
'cannot change after Seal, so the same pair always gives the same string,
'and Formulas.ParsedAnalysisFormula asks for it once per variable token of
'the same formula.
'@param tablename String. Structured reference table name.
'@param Connector Optional String. Connector between predicate fragments. Defaults to "*".
'@return String. Concatenated predicate or empty string when inputs are missing.
'@export
Public Function ConditionPredicate(ByVal tablename As String, Optional ByVal Connector As String = "*") As String
    Dim counter As Long
    Dim parsed As String
    Dim requestKey As String
    Dim vars As BetterArray
    Dim conds As BetterArray

    requestKey = tablename & KEY_SEPARATOR & Connector

    If this.cachedPredicateDone Then
        If StrComp(requestKey, this.cachedPredicateKey, vbBinaryCompare) = 0 Then
            ConditionPredicate = this.cachedPredicate
            Exit Function
        End If
    End If

    this.cachedPredicateDone = True
    this.cachedPredicateKey = requestKey
    this.cachedPredicate = vbNullString

    Set vars = this.varList
    Set conds = this.condList

    If vars.Length = 0 Or conds.Length = 0 Then Exit Function

    parsed = "(" & tablename & "[" & vars.Item(vars.LowerBound) & "]" & conds.Item(vars.LowerBound) & ")"

    For counter = vars.LowerBound + 1 To vars.UpperBound
        parsed = parsed & Connector & "(" & tablename & "[" & vars.Item(counter) & "]" & conds.Item(counter) & ")"
    Next

    this.cachedPredicate = parsed
    ConditionPredicate = parsed
End Function

'@section ErrorHandling
'===============================================================================
'@description Centralised helper to raise ProjectError-compliant errors.

'@label:throw-error
'@sub-title Raise a VBA error tagged with the class name
'@details
'Wrapper around Err.Raise that standardises the source to
'"FormulaCondition", providing a consistent stack trace.
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive text.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'@section PublicDiagnostics
'===============================================================================
'@description Defensive copies of the tracked lists and the diagnostic log.

'@label:variables-clone
'@prop-title Retrieve a clone of the stored variable names
'@details
'Returns a BetterArray copy of the variable identifiers in the order
'they were supplied at creation.
'@return BetterArray. The variable names preserving the original order.
'@export
Public Property Get Variables() As BetterArray
    Set Variables = this.varList.Clone()
End Property

'@label:conditions-clone
'@prop-title Retrieve a clone of the stored condition fragments
'@details
'Returns a BetterArray copy of the predicate fragments in the order
'they were supplied at creation.
'@return BetterArray. The condition fragments preserving the original order.
'@export
Public Property Get Conditions() As BetterArray
    Set Conditions = this.condList.Clone()
End Property

'@label:checking-values
'@prop-title Retrieve the internal checking log
'@details
'Returns the accumulated Checking instance containing all diagnostic
'messages from the last validation run, or Nothing when none recorded.
'@return Checking. The diagnostic log.
'@export
Public Property Get CheckingValues() As Checking
    If this.hasChecks Then Set CheckingValues = this.internalChecks
End Property
