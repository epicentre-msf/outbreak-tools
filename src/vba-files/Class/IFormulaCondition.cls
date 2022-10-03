VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of FormulaCondition"
'@Folder("Formulas")
'@ModuleDescription("Conditons Used to create formulas for analysis")
'@IgnoreModule

Option Explicit

Public Property Get Variable() As String: End Property
Public Property Get Condition() As String: End Property
Public Property Get Valid(ByVal Dict As IDictionary, ByVal tableName As String) As Boolean : End Property
Public Property Get ConditionString(ByVal tableName As String, ByVal varName As String, Optional Connector As String = "*") As String : End Property
