VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ImpSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Specifications for import in the disease worksheet"

'@Folder("Imports")
'@PredeclaredId
'@ModuleDescription("Specifications for import in the disease worksheet")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Implements IImpSpecs

Private Type TImp
  currwb As Workbook
  repfobj As Object
  needReport As Boolean
End Type

Private Const CLASSNAME  As String = "ImpSpecs"

