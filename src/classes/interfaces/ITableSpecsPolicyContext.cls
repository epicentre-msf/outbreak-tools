VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecsPolicyContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Context shared with TableSpecs table-type policies")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Accessors --------------------------------------------------------------------

'Return the analysis table type encoded as a byte (mirrors AnalysisTablesType).
Public Function TableTypeCode() As Byte
End Function

'Retrieve the stored value for a table specification column (case-insensitive).
Public Function ColumnValue(ByVal columnName As String) As String
End Function

'Return the configured row variable identifier.
Public Function RowVariable() As String
End Function

'Return the configured column variable identifier.
Public Function ColumnVariable() As String
End Function

'Variable interactions --------------------------------------------------------

'Check whether the variable collection contains the supplied variable.
Public Function VariablesContains(ByVal variableName As String) As Boolean
End Function

'Retrieve the control type associated to the provided variable.
Public Function VariableControl(ByVal variableName As String) As String
End Function

'Retrieve the variable type (e.g., "date") associated to the provided variable.
Public Function VariableType(ByVal variableName As String) As String
End Function

'Dictionary interactions ------------------------------------------------------

'Check whether the dictionary exposes a given variable name.
Public Function DictionaryVariableExists(ByVal variableName As String) As Boolean
End Function
