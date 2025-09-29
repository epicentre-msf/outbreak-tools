VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicyGlobalSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Implements TableSpecs rules for global summary tables")
'@depends TableSpecsPolicyHelpers
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableTypePolicy

Private Const CLASSNAME As String = "TableTypePolicyGlobalSummary"

'@section Helpers
'===============================================================================
'Fetch a table value once and normalise it for the comparison.
Private Function NormalizedValue(ByVal context As ITableSpecsPolicyContext, ByVal columnName As String) As String
    NormalizedValue = NormalizeValue(context.ColumnValue(columnName))
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = TABLE_TYPE_GLOBAL_SUMMARY
End Function

Private Function ITableTypePolicy_HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasPercent = False
End Function

Private Function ITableTypePolicy_HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasTotal = False
End Function

Private Function ITableTypePolicy_HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasMissing = False
End Function

Private Function ITableTypePolicy_HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasGraph = False
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_IsValid = HasText(NormalizedValue(context, "label")) And _
                               HasText(NormalizedValue(context, "function"))
End Function

