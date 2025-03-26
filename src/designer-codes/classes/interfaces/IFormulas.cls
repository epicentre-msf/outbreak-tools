VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Formulas Object"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Interface of Formulas Object")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods

Public Property Get VarLists() As BetterArray: End Property
Public Property Get FormulaTableList() As BetterArray: End Property

Public Property Get Valid(Optional ByVal formulaType As String = "analysis") As Boolean: End Property

Public Property Get ParsedLinelistFormula(Optional ByVal useTableName As Boolean = False, _
                                          Optional ByVal tablePrefix As String = vbNullString) As String
End Property

Public Property Get ParsedAnalysisFormula(ByVal formCond As IFormulaCondition, _
                                          Optional ByVal tablePrefix As String = vbNullString, _
                                          Optional ByVal Connector As String = "*") As String
End Property

'Test if there are litterals in the Formulas
Public Property Get HasLiterals() As Boolean: End Property
