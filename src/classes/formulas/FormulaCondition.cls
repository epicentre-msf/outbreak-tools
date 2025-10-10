VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Conditons Used to create formulas for analysis"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Conditions used to create formulas for analysis")
'@depends Checking
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation
'@section Overview
'===============================================================================
'@details Couples variables with their condition fragments, validates alignment, and produces structured Excel predicates for analysis formulas.
Option Explicit
Implements IFormulaCondition

'@section State
'===============================================================================
'@details Tracks the paired variable/condition lists alongside cached diagnostics.
Private Type TFormulaCondition
    condList As BetterArray                     'Conditions for the formulas
    varList As BetterArray                      'Variables for the formulas
    cachedTableName As String                    'Validated table name across variables
    hasChecks As Boolean
    internalChecks As IChecking
    checkCounter As Long
End Type


Private this As TFormulaCondition

'@section Factory
'===============================================================================
'@details Validate paired arrays and expose the ready-to-use condition helper.

'@label:Create
'@sub-title Construct a FormulaCondition instance from paired arrays
'@param formVar BetterArray containing variable names.
'@param formCond BetterArray containing predicate fragments.
'@return IFormulaCondition ready for validation/formatting.
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
'@details Provide controlled access to stored collections and interface pointer.

'@label:Self
'@sub-title Return the interface pointer for the current instance
'@return IFormulaCondition representing the current object.
Public Property Get Self() As IFormulaCondition
    Set Self = Me
End Property

'@label:LstConditions
'@sub-title Retrieve a clone of the stored condition fragments
'@return BetterArray preserving the original order.
Public Property Get LstConditions() As BetterArray
    Set LstConditions = this.condList
End Property

'@label:LstConditions
'@sub-title Store the condition fragments and reset caches
'@param formCond BetterArray containing predicate fragments.
Public Property Set LstConditions(ByVal formCond As BetterArray)
    Set this.condList = New BetterArray
    this.condList.LowerBound = 1
    this.condList.Items = formCond.Items
    ResetCaches
End Property

'@label:LstVars
'@sub-title Retrieve a clone of the stored variables list
'@return BetterArray preserving the original order.
Public Property Get LstVars() As BetterArray
    Set LstVars = this.varList
End Property

'@label:LstVars
'@sub-title Store the variables list, normalise bounds, and reset caches
'@param formVar BetterArray containing variable identifiers.
Public Property Set LstVars(ByVal formVar As BetterArray)
    Set this.varList = New BetterArray
    this.varList.LowerBound = 1
    this.varList.Items = formVar.Items
    ResetCaches
End Property

'@section InternalState
'===============================================================================
'@details Maintain cached validation artefacts and diagnostics.

'@label:ResetCaches
'@sub-title Clear cached table names and diagnostics information
Public Sub ResetCaches()
    this.cachedTableName = vbNullString
    this.hasChecks = False
    this.checkCounter = 0
    Set this.internalChecks = Nothing
End Sub

'@label:LogInfo
'@sub-title Append a diagnostic entry to the internal checking log
'@param message String describing the event.
Private Sub LogInfo(ByVal message As String, Optional ByVal scope As Byte = checkingInfo)
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create("FormulaCondition trace")
    this.internalChecks.Add CStr(this.checkCounter), message, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:HasInternalChecks
'@sub-title Determine whether diagnostics have been captured
'@return Boolean True when internal checks exist.
Private Property Get HasInternalChecks() As Boolean
    HasInternalChecks = this.hasChecks
End Property

'@label:CachedTable
'@sub-title Retrieve the memoised table name resolved during validation
'@return String cached table identifier.
Private Property Get CachedTable() As String
    CachedTable = this.cachedTableName
End Property

'@section MetadataHelpers
'===============================================================================
'@details Provide metadata derived from the last validation run.

'@label:VariablesTable
'@sub-title Resolve the table name shared by all tracked variables
'@param dict ILLdictionary used when validation has not occurred yet.
'@return String table identifier or empty string when unavailable.
Public Property Get VariablesTable(ByVal dict As ILLdictionary) As String
    If LenB(CachedTable) = 0 Then
        If Not Valid(dict) Then Exit Property
    End If
    VariablesTable = CachedTable
End Property

'@section Validation
'===============================================================================
'@details Confirm that variables/conditions are aligned and belong to the same table.

'@label:Valid
'@sub-title Validate variable/condition lists against the dictionary
'@param dict ILLdictionary metadata provider.
'@param tablename Optional String forcing the expected table.
'@return Boolean True when validation succeeds.
Public Function Valid(ByVal dict As ILLdictionary, Optional ByVal tablename As String = vbNullString) As Boolean

    Dim vars As ILLVariables
    Dim varList As BetterArray
    Dim condList As BetterArray
    Dim idx As Long
    Dim firstVar As String
    Dim targetTable As String
    Dim currentTable As String

    ResetCaches
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

    For idx = varList.LowerBound To varList.UpperBound
        Dim varName As String
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
    Valid = True
    Exit Function

ValidationError:
    LogInfo "Validation failed: " & Err.Description
    Err.Clear
End Function

'@section FormulaRendering
'===============================================================================
'@details Produce Excel-ready predicates and wrapper formulas.

'@label:ConditionString
'@sub-title Wrap the predicate in an IF expression referencing the supplied variable
'@param tablename String structured reference table name.
'@param varName String variable/column name.
'@param Connector Optional String connector used between predicates.
'@return String IF expression or empty string when predicate is empty.
Public Function ConditionString(ByVal tablename As String, ByVal varName As String, _
                                Optional ByVal Connector As String = "*") As String

    Dim predicate As String

    predicate = ConditionPredicate(tablename, Connector)
    If LenB(predicate) = 0 Then Exit Function

    ConditionString = "IF(" & predicate & " , " & tablename & "[" & varName & "]" & ")"
End Function

'@label:ConditionPredicate
'@sub-title Join the variable predicates using the supplied connector
'@param tablename String structured reference table name.
'@param Connector Optional String connector between predicate fragments.
'@return String concatenated predicate or empty string when inputs are missing.
Public Function ConditionPredicate(ByVal tablename As String, Optional ByVal Connector As String = "*") As String
    Dim counter As Long
    Dim parsed As String

    If Variable.Length = 0 Or Condition.Length = 0 Then Exit Function

    parsed = "(" & tablename & "[" & Variable.Item(Variable.LowerBound) & "]" & Condition.Item(Variable.LowerBound) & ")"

    For counter = Variable.LowerBound + 1 To Variable.UpperBound
        parsed = parsed & Connector & "(" & tablename & "[" & Variable.Item(counter) & "]" & Condition.Item(counter) & ")"
    Next

    ConditionPredicate = parsed
End Function

'@section ErrorHandling
'===============================================================================
'@details Centralised helper to raise ProjectError-compliant errors.

'@label:ThrowError
'@sub-title Raise a VBA error tagged with the class name
'@param errNumb Long ProjectError code.
'@param errorMessage String descriptive text.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, "FormulaCondition", errorMessage
End Sub

'@section Interface
'===============================================================================
'@details Explicit interface implementations delegating to the public members.
Private Property Get IFormulaCondition_Variables() As BetterArray
    Set IFormulaCondition_Variables = Variable
End Property

Private Property Get IFormulaCondition_Conditions() As BetterArray
    Set IFormulaCondition_Conditions = Condition
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

Private Property Get IFormulaCondition_CheckingValues() As IChecking
    If HasInternalChecks Then Set IFormulaCondition_CheckingValues = this.internalChecks
End Property
