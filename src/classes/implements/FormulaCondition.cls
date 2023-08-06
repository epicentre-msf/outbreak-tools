VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A Class for dealing with Conditions in formulas"

'@Folder("Formulas")
'@ModuleDescription("Conditons Used to create formulas for analysis")
'@IgnoreModule
'@PredeclareId
Option Explicit
Implements IFormulaCondition

'Types of the class
Private Type TFormulaCondition
    Condition As BetterArray                     'Conditions for the formulas
    Variable As BetterArray                      'Variables for the formulas
End Type

Private Const CLASSNAME As String = "Formula"
Private this As TFormulaCondition

'Instanciation
Public Function Create(ByVal formVar As BetterArray, ByVal formCond As BetterArray) As IFormulaCondition

    'Seems like BetterArry is referencing the Arrays within my object. I want them to be
    'immutable, so I clone those arrays.
    
    If formVar.Length <> formCond.Length Then ThrowErrorUnableInstanciate "Condition length: " & formCond.Length & "different from variable length: " & formVar.Length

    With New FormulaCondition
        Set .Condition = formCond
        Set .Variable = formVar
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
End Property

'Variables
Public Property Get Variable() As BetterArray
    Set Variable = this.Variable.Clone()
End Property

Public Property Set Variable(ByVal formVar As BetterArray)
    Set this.Variable = formVar.Clone()
End Property

Public Property Get VariablesTable(ByVal dict As ILLdictionary) As String
    Dim vars As ILLVariables
    Dim firstVar As String
    If Valid(dict) Then
        Set vars = LLVariables.Create(dict)
        firstVar = Variable.Item(Variable.LowerBound)
        VariablesTable = vars.Value(varName:=firstVar, colName:="table name")
    End If
End Property

'Test if a formula is valid:
Public Function Valid(ByVal dict As ILLdictionary, Optional ByVal tableName As String = vbNullString) As Boolean

    Dim tableData As BetterArray                 'list of all tables in dictionary
    Dim variableData As BetterArray              'list of all variables in dictionary
    Dim formulavariables As BetterArray
    Dim formulaConditions As BetterArray
    Dim customtableName As String

    Dim variablesTableName As String
    Dim VARIABLENAME As String
    Dim counter As Long

    Set tableData = New BetterArray
    Set variableData = New BetterArray
    Set formulavariables = New BetterArray
    Set formulaConditions = New BetterArray

    Set formulavariables = Variable
    Set formulaConditions = Condition

    'Valid Fail by default, turns to true after the checkings
    Valid = False

    If Not (dict.ColumnExists("table name") And dict.ColumnExists("variable name")) Then Exit Function

    variableData.FromExcelRange dict.DataRange("variable name")
    tableData.FromExcelRange dict.DataRange("table name")


    'variable and conditions table doesn't fit
    If formulavariables.Length <> formulaConditions.Length Or formulavariables.Length = 0 Then Exit Function

    'test the table name
    counter = formulavariables.LowerBound
    If tableName = vbNullString Then
        VARIABLENAME = formulavariables.Item(counter)
        If Not variableData.Includes(VARIABLENAME) Then Exit Function
        customtableName = tableData.Item(variableData.IndexOf(VARIABLENAME))
    Else
        customtableName = tableName
    End If


    For counter = formulavariables.LowerBound To formulavariables.UpperBound

        VARIABLENAME = formulavariables.Item(counter)

        'Unable to find the variable within the provided list of variables
        If Not variableData.Includes(VARIABLENAME) Then Exit Function

        'Variables are not in the same table
        variablesTableName = tableData.Item(variableData.IndexOf(VARIABLENAME))
        If variablesTableName <> customtableName Then Exit Function
    Next

    Valid = True
End Function

'Convert the conditions into a string of formulas
Public Function ConditionString(ByVal tableName As String, ByVal varName As String, _
                                Optional ByVal Connector As String = "*") As String

    'The connector is how we want to link the conditions: * for AND, + for OR

    Dim counter As Long
    Dim ParsedCondition As String

    ParsedCondition = vbNullString
    ConditionString = vbNullString

    'You get an empty vector in case of no data
    If Variable.Length = 0 Or Condition.Length = 0 Then Exit Function

    'Initialize
    counter = Variable.LowerBound
    ParsedCondition = "(" & tableName & "[" & Variable.Item(counter) & "]" & Condition.Item(counter) & ")"

    'parse every condition
    For counter = Variable.LowerBound + 1 To Variable.UpperBound
        ParsedCondition = ParsedCondition & Connector & "(" & tableName & "[" & Variable.Item(counter) & "]" & Condition.Item(counter) & ")"
    Next

    ConditionString = "IF(" & ParsedCondition & " , " & tableName & "[" & varName & "]" & ")"
End Function

'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Implement interfaces
Public Property Get IFormulaCondition_Variables() As BetterArray
    Set IFormulaCondition_Variables = Variable
End Property

Public Property Get IFormulaCondition_Conditions() As BetterArray
    Set IFormulaCondition_Conditions = Condition
End Property

Public Property Get IFormulaCondition_VariablesTable(ByVal dict As ILLdictionary) As String
    IFormulaCondition_VariablesTable = VariablesTable(dict)
End Property

Public Function IFormulaCondition_ConditionString(ByVal tableName As String, ByVal varName As String, _
                                                  Optional ByVal Connector As String = "*") As String
    IFormulaCondition_ConditionString = ConditionString(tableName:=tableName, varName:=varName, Connector:=Connector)
End Function

Public Function IFormulaCondition_Valid(ByVal dict As ILLdictionary, Optional ByVal tableName As String = vbNullString) As Boolean
    IFormulaCondition_Valid = Valid(dict, tableName)
End Function
