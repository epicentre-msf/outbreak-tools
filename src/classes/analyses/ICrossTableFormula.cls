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

'@class ICrossTableFormula
'@description
'Defines the public contract for the CrossTableFormula class. Exposes a
'validity check and a method to write all computed formulas into the
'cross-table cells on the analysis output worksheet.


' Formula scope for linelist vs analysis formulas
Public Enum ExcelFormulaScope
    FormulaScopeLinelist = 1
    FormulaScopeAnalysis = 2
End Enum


'@section Properties
'===============================================================================

'@jump:valid
'@prop-title Whether the formula object is valid for the analysis context
'@details
'Returns True when the underlying formula definition and cross-table
'reference are both properly initialised. Callers should check this
'before calling AddFormulas to avoid writing incomplete output.
'@return Boolean. True when the formula object is ready to write.
Public Property Get Valid() As Boolean: End Property


'@section Methods
'===============================================================================

'@jump:add-formulas
'@sub-title Write all formulas to the cross-table on the worksheet
'@details
'Iterates through the cross-table cells and writes the appropriate Excel
'formula into each data cell. Handles column formulas, row formulas,
'percentage formulas, and total formulas based on the table scope and
'specification flags. Skips writing when Valid returns False.
'@export
Public Sub AddFormulas(): End Sub
