VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of FormulaData"

'@Interface
'@Folder("Formulas")
'@ModuleDescription("Interface of FormulaData")

Option Explicit

'Exposed methods
Public Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean: End Function
Public Function ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean: End Function

