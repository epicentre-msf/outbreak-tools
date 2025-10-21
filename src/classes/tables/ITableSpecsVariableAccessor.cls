VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableSpecsVariableAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Abstraction over linelist variables used by TableSpecs policies")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Check whether the dictionary contains the supplied variable (case-sensitive by default).
Public Function Contains(ByVal variableName As String) As Boolean
End Function

'Retrieve the control type associated to a variable.
Public Function Control(ByVal variableName As String) As String
End Function

'Retrieve the recorded variable type (e.g., "date", "text").
Public Function VariableType(ByVal variableName As String) As String
End Function
