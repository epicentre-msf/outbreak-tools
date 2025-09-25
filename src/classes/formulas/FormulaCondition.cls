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
'@IgnoreModule UseMeaningfulName, HungarianNotation
Option Explicit
Implements IFormulaCondition

'Types of the class
Private Type TFormulaCondition
    Condition As BetterArray                     'Conditions for the formulas
    Variable As BetterArray                      'Variables for the formulas
    cachedTableName As String                    'Validated table name across variables
    hasChecks As Boolean
    internalChecks As IChecking
    checkCounter As Long
End Type

Private Const CLASSNAME As String = "FormulaCondition"
Private this As TFormulaCondition

'Instanciation
Public Function Create(ByVal formVar As BetterArray, ByVal formCond As BetterArray) As IFormulaCondition

    If formVar Is Nothing Then ThrowError ProjectError.InvalidArgument, "Variables array cannot be null"
    If formCond Is Nothing Then ThrowError ProjectError.InvalidArgument, "Conditions array cannot be null"

    If formVar.Length <> formCond.Length Then _
        ThrowError ProjectError.InvalidArgument, _
                  "Condition length " & formCond.Length & " different from variable length " & formVar.Length

    With New FormulaCondition
        Set .Condition = formCond
        Set .Variable = formVar
        .ResetCaches
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IFormulaCondition
    Set Self = Me
End Property

'Conditions
Public Property Get Condition() As BetterArray
    Set Condition = this.Condition.Clone()
End Property

Public Property Set Condition(ByVal formCond As BetterArray)
    Set this.Condition = formCond.Clone()
    this.Condition.LowerBound = 1
    ResetCaches
End Property

'Variables
Public Property Get Variable() As BetterArray
    Set Variable = this.Variable.Clone()
End Property

Public Property Set Variable(ByVal formVar As BetterArray)
    Set this.Variable = formVar.Clone()
    this.Variable.LowerBound = 1
    ResetCaches
End Property

'@section InternalState
'===============================================================================
Private Sub ResetCaches()
    this.cachedTableName = vbNullString
    this.hasChecks = False
    this.checkCounter = 0
    Set this.internalChecks = Nothing
End Sub

Private Sub LogInfo(ByVal message As String)
    If this.internalChecks Is Nothing Then Set this.internalChecks = Checking.Create("FormulaCondition trace")
    this.internalChecks.Add CStr(this.checkCounter), message
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

Private Property Get HasInternalChecks() As Boolean
    HasInternalChecks = this.hasChecks
End Property

Private Property Get CachedTable() As String
    CachedTable = this.cachedTableName
End Property

'@section MetadataHelpers
'===============================================================================
Public Property Get VariablesTable(ByVal dict As ILLdictionary) As String
    If LenB(CachedTable) = 0 Then
        If Not Valid(dict) Then Exit Property
    End If
    VariablesTable = CachedTable
End Property

'Test if a formula is valid:
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

    Set varList = Variable
    Set condList = Condition

    If varList.Length = 0 Or condList.Length = 0 Then
        LogInfo "Validation failed: empty variable or condition list"
        Exit Function
    End If

    If varList.Length <> condList.Length Then
        LogInfo "Validation failed: variable/condition length mismatch"
        Exit Function
    End If

    On Error GoTo ValidationError
        Set vars = LLVariables.Create(dict)
    On Error GoTo 0

    firstVar = CStr(varList.Item(varList.LowerBound))
    If Not vars.Contains(firstVar) Then
        LogInfo "Validation failed: variable not found - " & firstVar
        Exit Function
    End If

    If tablename = vbNullString Then
        targetTable = vars.TableName(firstVar)
    Else
        targetTable = tablename
    End If

    If LenB(targetTable) = 0 Then
        LogInfo "Validation failed: table name not resolved"
        Exit Function
    End If

    For idx = varList.LowerBound To varList.UpperBound
        Dim varName As String
        varName = CStr(varList.Item(idx))

        If Not vars.Contains(varName) Then
            LogInfo "Validation failed: variable not found - " & varName
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

"Convert the conditions into a string of formulas
Public Function ConditionString(ByVal tablename As String, ByVal varName As String, _
                                Optional ByVal Connector As String = "*") As String

    Dim predicate As String

    predicate = ConditionPredicate(tablename, Connector)
    If LenB(predicate) = 0 Then Exit Function

    ConditionString = "IF(" & predicate & " , " & tablename & "[" & varName & "]" & ")"
End Function

'@label:ConditionPredicate
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

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowError(ByVal errNumb As ProjectError, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, errorMessage
End Sub

'Implement interfaces
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
