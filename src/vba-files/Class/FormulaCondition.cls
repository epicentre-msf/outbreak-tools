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
    Condition As BetterArray 'Conditions for the formulas
    Variable As BetterArray  'Variables for the formulas
End Type


Private this As TFormulaCondition

'Instanciation
Public Function Create(ByVal formCond As BetterArray, ByVal formVar As BetterArray) As IFormulaCondition
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
    Set Condition = this.Condition
End Property

Public Property Set Condition(ByVal formCond As BetterArray)
    this.Condition = formCond
End Property

'Variables
Public Property Get Variable() As BetterArray
    Set Variable = this.Variable
End Property

Public Property Let Variable(ByVal formVar As BetterArray)
     Set this.Variable = formVar
End Property

'Test if a formula is valid:
Public Property Get Valid(ByVal Dict As ILLdictionary, ByVal tableName As String) As Boolean

    Dim tableData As BetterArray 'list of all tables in dictionary
    Dim variableData As BetterArray 'list of all variables in dictionary

    Dim variablesTableName As String
    Dim variableName As String
    Dim counter As Long

    'Valid Fail by default, turns to true after the checkings
    Valid = False

    'variable and conditions table doesn't fit
    If Variable.Length <> Condition.Length Or Variable.Length = 0 Then Exit Property

    Set tableData = Dict.Column("table name")
    Set variableData = Dict.Column("variable name")

    For counter = Variable.LowerBound To Variable.UpperBound

        variableName = Variable.Item(counter)

        'Unable to find the variable within the provided list of variables
        If Not variableData.Includes(variableName) Then Exit Property

        'Variables are not in the same table
        variablesTableName = tableData.Item(variableData.IndexOf(variableName))
        If variablesTableName <> tableName Then Exit Property
    Next

    Valid = True
End Property


'Convert the conditions into a string of formulas
Public Property Get ConditionString(ByVal tableName As String, ByVal varName As String, Optional Connector As String = "*") As String

    'The connector is how we want to link the conditions: * for AND, + for OR

    Dim counter As Long
    Dim ParsedCondition As String

    ParsedCondition = vbNullString
    ConditionString = vbNullString

    'You get an empty vector in case of no data
    If Variable.Length = 0 Or Condition.Length = 0 Then Exit Property

    'Initialize
    counter = Variable.LowerBound
    ParsedCondition = "(" & tableName & "[" & Variable.Item(counter) & "]" & Condition.Item(counter) & ")"

    'parse every condition
    For counter = Variable.LowerBound + 1 To Variable.UpperBound
        ParsedCondition = ParsedCondition & Connector & "(" & tableName & "[" & Variable.Item(counter) & "]" & Condition.Item(counter) & ")"
    Next

    ConditionString = "IF(" & ParsedCondition & "," & tableName & "[" & varName & "]" & ")"
End Property

'Implement interfaces
Public Property Get IFormulaCondition_Variable() As BetterArray
    Set IFormulaCondition_Variable = Variable
End Property

Public Property Get IFormulaCondition_Condition() As BetterArray
     Set IFormulaCondition_Condition = Condition
End Property

Public Property Get IFormulaCondition_ConditionString(ByVal tableName As String, ByVal varName As String, Optional Connector As String = "*") As String
    IFormulaCondition_ConditionString = ConditionString(tableName:=tableName, varName:=varName, Connector:=Connector)
End Property

Public Property Get IFormulaCondition_Valid(ByVal Dict As IDictionary, ByVal tableName As String) As Boolean
    IFormulaCondition_Valid = Valid(Dict, tableName)
End Property
