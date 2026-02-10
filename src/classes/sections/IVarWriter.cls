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

'@class IVarWriter
'@description
'Defines the public contract for the VarWriter class, exposing methods
'for writing a single dictionary variable to one or more linelist
'worksheets and for reading dictionary column values. Consumers call
'WriteVariable once per variable; the implementation handles main,
'printed-companion, and CRF-companion sheets according to the active
'layer (HList or VList).

'@details
'Layers control which worksheets are written. HList writes to the main
'sheet plus optional printed and CRF companions. VList writes to the
'main sheet only.
Public Enum VarWriterLayer
    VarWriterLayerHList = 1
    VarWriterLayerVList = 2
End Enum


'@section Methods
'===============================================================================

'@jump:write-variable
'@sub-title Write one variable to the target worksheet(s)
'@details
'Writes label, formula, type formatting, choices, validation, and
'conditional formatting for the named variable. On HList layers the
'method also writes to the printed and CRF companion sheets when they
'are available. Internally resets cached cell ranges before each
'invocation so the writer can be reused across multiple variables.
'@param varName String. The variable name as it appears in the dictionary.
'@export
Public Sub WriteVariable(ByVal varName As String)
End Sub

'@jump:value-of
'@prop-title Read a dictionary column value for the current variable
'@details
'Returns the value stored in the specified dictionary column for the
'variable that was most recently passed to WriteVariable. Delegates to
'LLVariables.Value internally.
'@param colName String. The dictionary column name to look up.
'@return String. The column value for the current variable.
'@export
Public Property Get ValueOf(ByVal colName As String) As String
End Property
