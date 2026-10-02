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
Implements IFormulaCondition

'@class FormulaCondition
'@description
'Couples variable names with their condition fragments, validates alignment
'against the dictionary, and produces structured Excel predicates for
'analysis formulas. The class pairs two BetterArray lists (variables and
'conditions), validates that all variables belong to the same table, and
'exposes methods to render IF-wrapped condition strings and joined predicates.
'Consumers interact through the IFormulaCondition interface.
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
    internalChecks As IChecking
    checkCounter As Long
    cachedValidDone As Boolean                   'Whether Valid has produced a cached result
    cachedValidResult As Boolean                 'Cached True/False from last Valid call
    cachedValidTable As String                   'The tablename parameter used for cached result
End Type


Private this As TFormulaCondition

'@section Factory
'===============================================================================
'@description Validate paired arrays and expose the ready-to-use condition helper.

'@label:create
'@sub-title Construct a FormulaCondition instance from paired arrays
'@details
'Validates that both arrays are non-Nothing and have equal lengths, then
'stores normalised copies with a lower bound of 1. The returned instance
'is cast to IFormulaCondition for consistent downstream consumption.
'@param formVar BetterArray. The array containing variable names.
'@param formCond BetterArray. The array containing predicate fragments.
'@return IFormulaCondition. A ready-to-use condition helper.
'@throws ProjectError.InvalidArgument When formVar or formCond is Nothing or lengths differ.
'@export
Public Function Create(ByVal formVar As BetterArray, ByVal formCond As BetterArray) As IFormulaCondition

    If formVar Is Nothing Then ThrowError ProjectError.InvalidArgument, "Variables array cannot be null"
    If formCond Is Nothing Then ThrowError ProjectError.InvalidArgument, "Conditions array cannot be null"

    If formVar.Length <> formCond.Length Then _
        ThrowError ProjectError.InvalidArgument, _
                  "Condition length " & formCond.Length & " different from variable length " & formVar.Length

    With New FormulaCondition
        Set .LstConditions = formCond
        Set .LstVars = formVar
        Set Create = .Self
    End With
End Function

'@section PublicAccessors
'===============================================================================
'@description Provide controlled access to stored collections and interface pointer.

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to its interface. Used internally
'by the factory method to return an IFormulaCondition reference.
'@return IFormulaCondition. The current object.
Public Property Get Self() As IFormulaCondition
    Set Self = Me
End Property

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
'@param formCond BetterArray. The array containing predicate fragments.
Public Property Set LstConditions(ByVal formCond As BetterArray)
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
'Resets all memoised validation state so the next call to Valid or
'VariablesTable triggers a fresh evaluation. Also clears any
'diagnostic entries accumulated during previous validations.
Public Sub ResetCaches()
    this.cachedTableName = vbNullString
    this.hasChecks = False
    this.checkCounter = 0
    Set this.internalChecks = Nothing
    this.cachedValidDone = False
    this.cachedValidResult = False
    this.cachedValidTable = vbNullString
End Sub

'@label:log-info
'@sub-title Append a diagnostic entry to the internal checking log
'@details
'Creates the checking instance on first use and appends the message
'using a monotonically increasing counter key. The checking store
'is exposed through the interface for callers needing diagnostic traces.
'@param message String. The event description to record.
'@param scope Optional Byte. Severity level from CheckingLogType. Defaults to checkingInfo.
Private Sub LogInfo(ByVal message As String, Optional ByVal scope As Byte = checkingInfo)
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create("FormulaCondition trace")
    this.internalChecks.Add CStr(this.checkCounter), message, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:has-internal-checks
'@prop-title Determine whether diagnostics have been captured
'@details
'Returns True when at least one LogInfo call has been made since the
'last ResetCaches, indicating that diagnostic information is available.
'@return Boolean. True when internal checks exist.
'@export
Private Property Get HasInternalChecks() As Boolean
    HasInternalChecks = this.hasChecks
End Property

'@label:cached-table
'@prop-title Retrieve the memoised table name resolved during validation
'@details
'Holds the table name that all variables were confirmed to share
'after a successful call to Valid. Empty until validation completes.
'@return String. Cached table identifier.
Private Property Get CachedTable() As String
    CachedTable = this.cachedTableName
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
'@param dict ILLdictionary. The dictionary used when validation has not occurred yet.
'@return String. Table identifier or empty string when unavailable.
'@export
Public Property Get VariablesTable(ByVal dict As ILLdictionary) As String
    If LenB(CachedTable) = 0 Then
        If Not Valid(dict) Then Exit Property
    End If
    VariablesTable = CachedTable
End Property

'@section Validation
'===============================================================================
'@description Confirm that variables/conditions are aligned and belong to the same table.

'@label:valid
'@sub-title Validate variable/condition lists against the dictionary
'@details
'Checks that both lists are non-empty and equal in length, then
'verifies each variable exists in the dictionary and belongs to the
'expected table. When a tablename argument is supplied it overrides
'the table resolved from the first variable. Results are cached so
'repeated calls with the same tablename parameter return immediately
'without re-validating. The cache is invalidated when underlying data
'changes (via Set LstVars or Set LstConditions).
'@param dict ILLdictionary. The metadata provider used for variable lookups.
'@param tablename Optional String. When supplied, forces the expected table name. Defaults to vbNullString.
'@return Boolean. True when validation succeeds.
'@export
Public Function Valid(ByVal dict As ILLdictionary, Optional ByVal tablename As String = vbNullString) As Boolean

    Dim vars As ILLVariables
    Dim varList As BetterArray
    Dim condList As BetterArray
    Dim idx As Long
    Dim firstVar As String
    Dim targetTable As String
    Dim currentTable As String

    ' Return cached result when validation was already done with matching parameters.
    ' The cache is cleared by ResetCaches (called from Set LstVars / Set LstConditions)
    ' whenever the underlying data changes.
    If this.cachedValidDone Then
        If StrComp(tablename, this.cachedValidTable, vbBinaryCompare) = 0 Then
            Valid = this.cachedValidResult
            Exit Function
        End If
    End If

    ' Fresh validation needed — clear prior diagnostics
    this.cachedTableName = vbNullString
    this.hasChecks = False
    this.checkCounter = 0
    Set this.internalChecks = Nothing

    ' Mark cache as populated (default result is False)
    this.cachedValidDone = True
    this.cachedValidTable = tablename
    this.cachedValidResult = False
    Valid = False

    Set varList = LstVars()
    Set condList = LstConditions()

    If varList.Length = 0 Or condList.Length = 0 Then
        LogInfo "Validation failed: empty variable or condition list", checkingError
        Exit Function
    End If

    If varList.Length <> condList.Length Then
        LogInfo "Validation failed: variable/condition length mismatch", checkingError
        Exit Function
    End If

    On Error GoTo ValidationError
        Set vars = LLVariables.Create(dict)
    On Error GoTo 0

    firstVar = CStr(varList.Item(varList.LowerBound))
    If Not vars.Contains(firstVar) Then
        LogInfo "Validation failed: variable not found - " & firstVar, checkingError
        Exit Function
    End If

    If tablename = vbNullString Then
        targetTable = vars.TableName(firstVar)
    Else
        targetTable = tablename
    End If

    If LenB(targetTable) = 0 Then
        LogInfo "Validation failed: table name not resolved", checkingError
        Exit Function
    End If

    Dim varName As String
    For idx = varList.LowerBound To varList.UpperBound

        varName = CStr(varList.Item(idx))

        If Not vars.Contains(varName) Then
            LogInfo "Validation failed: variable not found - " & varName, checkingError
            Exit Function
        End If

        currentTable = vars.TableName(varName)
        If StrComp(currentTable, targetTable, vbBinaryCompare) <> 0 Then
            LogInfo "Validation failed: variable belongs to different table - " & varName
            Exit Function
        End If
    Next idx

    this.cachedTableName = targetTable
    this.cachedValidResult = True
    Valid = True
    Exit Function

ValidationError:
    LogInfo "Validation failed: " & Err.Description
    Err.Clear
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
'structured table references.
'@param tablename String. Structured reference table name.
'@param Connector Optional String. Connector between predicate fragments. Defaults to "*".
'@return String. Concatenated predicate or empty string when inputs are missing.
'@export
Public Function ConditionPredicate(ByVal tablename As String, Optional ByVal Connector As String = "*") As String
    Dim counter As Long
    Dim parsed As String

    If LstVars.Length = 0 Or LstConditions.Length = 0 Then Exit Function

    parsed = "(" & tablename & "[" & LstVars.Item(LstVars.LowerBound) & "]" & LstConditions.Item(LstVars.LowerBound) & ")"

    For counter = LstVars.LowerBound + 1 To LstVars.UpperBound
        parsed = parsed & Connector & "(" & tablename & "[" & LstVars.Item(counter) & "]" & LstConditions.Item(counter) & ")"
    Next

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
    Err.Raise errNumb, "FormulaCondition", errorMessage
End Sub

'@section Interface
'===============================================================================
'@description Explicit interface implementations delegating to the public members.

Private Property Get IFormulaCondition_Variables() As BetterArray
    Set IFormulaCondition_Variables = LstVars.Clone()
End Property

Private Property Get IFormulaCondition_Conditions() As BetterArray
    Set IFormulaCondition_Conditions = LstConditions.Clone()
End Property

Private Property Get IFormulaCondition_VariablesTable(ByVal dict As ILLdictionary) As String
    IFormulaCondition_VariablesTable = VariablesTable(dict)
End Property

Private Function IFormulaCondition_ConditionString(ByVal tablename As String, ByVal varName As String, _
                                                  Optional ByVal Connector As String = "*") As String
    IFormulaCondition_ConditionString = ConditionString(tablename:=tablename, varName:=varName, Connector:=Connector)
End Function

Private Function IFormulaCondition_ConditionPredicate(ByVal tablename As String, Optional ByVal Connector As String = "*") As String
    IFormulaCondition_ConditionPredicate = ConditionPredicate(tablename, Connector)
End Function

Private Function IFormulaCondition_Valid(ByVal dict As ILLdictionary, Optional ByVal tablename As String = vbNullString) As Boolean
    IFormulaCondition_Valid = Valid(dict, tablename)
End Function

Private Function IFormulaCondition_HasCheckings() As Boolean
    IFormulaCondition_HasCheckings = HasInternalChecks
End Function

'@label:checking-values
Private Property Get IFormulaCondition_CheckingValues() As IChecking
    If HasInternalChecks Then Set IFormulaCondition_CheckingValues = this.internalChecks
End Property
