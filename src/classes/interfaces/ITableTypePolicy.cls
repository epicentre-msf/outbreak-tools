VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITableTypePolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Defines behaviour for a specific analysis table type")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'Returns the table type the policy represents.
Public Function TableTypeCode() As Byte
End Function

'Determine whether the table exposes percentage output.
Public Function HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
End Function

'Determine whether the table exposes totals.
Public Function HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
End Function

'Determine whether the table exposes missing values.
Public Function HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
End Function

'Determine whether a graph should be generated for the table.
Public Function HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
End Function

'Validate whether the table specification is complete.
Public Function IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
End Function
