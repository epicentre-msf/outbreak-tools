VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVarWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Sections")
'@ModuleDescription("Interface for writing one variable to a linelist worksheet")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

'@description Layers for variable writing. HList writes main + printed + CRF companion.
'   VList writes main sheet only.
Public Enum VarWriterLayer
    VarWriterLayerHList = 1
    VarWriterLayerVList = 2
End Enum

'@section Core Methods
'===============================================================================

'@description Write one variable to the target worksheet(s).
'@param varName The variable name as it appears in the dictionary.
Public Sub WriteVariable(ByVal varName As String)
End Sub

'@description Read a dictionary column value for the current variable.
'@param colName The dictionary column name.
'@return The value as a string.
Public Property Get ValueOf(ByVal colName As String) As String
End Property
