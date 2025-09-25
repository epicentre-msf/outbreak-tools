VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicyUnivariate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Implements TableSpecs rules for univariate tables")
'@depends TableSpecsPolicyHelpers
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableTypePolicy

Private Const CLASSNAME As String = "TableTypePolicyUnivariate"

'@section Helpers
'===============================================================================
'Normalise a column value for easier comparison.
Private Function NormalizedValue(ByVal context As ITableSpecsPolicyContext, ByVal columnName As String) As String
    NormalizedValue = NormalizeValue(context.ColumnValue(columnName))
End Function

'Determine whether the supplied control is supported by the analysis rules.
Private Function IsChoiceControl(ByVal controlValue As String) As Boolean
    IsChoiceControl = ValueInList(controlValue, "choice_manual", "choice_formula")
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = TABLE_TYPE_UNIVARIATE
End Function

Private Function ITableTypePolicy_HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasPercent = ValueEquals(context.ColumnValue("percentage"), "yes")
End Function

Private Function ITableTypePolicy_HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasTotal = True
End Function

Private Function ITableTypePolicy_HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasMissing = ValueEquals(context.ColumnValue("missing"), "yes")
End Function

Private Function ITableTypePolicy_HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasGraph = ValueEquals(context.ColumnValue("graph"), "yes")
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    Dim rowVariable As String
    Dim rowControl As String

    rowVariable = context.RowVariable
    If Not context.VariablesContains(rowVariable) Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    rowControl = context.VariableControl(rowVariable)
    ITableTypePolicy_IsValid = IsChoiceControl(rowControl)
End Function

