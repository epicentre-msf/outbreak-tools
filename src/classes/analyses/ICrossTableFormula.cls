VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICrossTableFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for cross-table formula writing"

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface for cross-table formula writing")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

' Formula scope for linelist vs analysis formulas
Public Enum ExcelFormulaScope
    FormulaScopeLinelist = 1
    FormulaScopeAnalysis = 2
End Enum

' @description Whether the formula object is valid for analysis context
Public Property Get Valid() As Boolean: End Property

' @description Write all formulas to the cross-table on the worksheet
Public Sub AddFormulas(): End Sub
