VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicySpatioTemporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Implements TableSpecs rules for spatio-temporal tables")
'@depends TableSpecsPolicyHelpers
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableTypePolicy

Private Const CLASSNAME As String = "TableTypePolicySpatioTemporal"

'@section Helpers
'===============================================================================
'Check whether the dictionary exposes any of the expected spatial prefixes.
Private Function HasSpatialVariable(ByVal context As ITableSpecsPolicyContext, ByVal baseName As String) As Boolean
    HasSpatialVariable = context.DictionaryVariableExists("adm1_" & baseName) Or _
                         context.DictionaryVariableExists("hf_" & baseName)
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = TABLE_TYPE_SPATIO_TEMPORAL
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
    ITableTypePolicy_HasGraph = ValueEquals(context.ColumnValue("graph"), "yes")
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    Dim rowVariable As String
    Dim columnVariable As String

    rowVariable = context.RowVariable
    columnVariable = context.ColumnVariable

    If Not ValueEquals(context.VariableType(rowVariable), "date") Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    ITableTypePolicy_IsValid = HasSpatialVariable(context, columnVariable)
End Function

