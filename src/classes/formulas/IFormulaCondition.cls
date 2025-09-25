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
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Public Property Get Variables() As BetterArray: End Property
Public Property Get Conditions() As BetterArray: End Property

Public Property Get VariablesTable(ByVal dict As ILLdictionary) As String: End Property

'Methods
Public Function Valid(ByVal dict As ILLdictionary, Optional ByVal tablename As String = vbNullString) As Boolean: End Function

Public Function ConditionString(ByVal tablename As String, ByVal varName As String, _
                                Optional ByVal Connector As String = "*") As String
End Function

'Return the joined predicate without IF wrapper.
Public Function ConditionPredicate(ByVal tablename As String, Optional ByVal Connector As String = "*") As String
End Function

'Exposure of validation traces for diagnostics.
Public Function HasCheckings() As Boolean
End Function

Public Property Get CheckingValues() As IChecking
End Property
