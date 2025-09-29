VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicyBivariate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Implements TableSpecs rules for bivariate tables")
'@depends TableSpecsPolicyHelpers
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableTypePolicy

Private Const CLASSNAME As String = "TableTypePolicyBivariate"

'@section Helpers
'===============================================================================
'Determine whether the supplied control is supported by the analysis rules.
Private Function IsChoiceControl(ByVal controlValue As String) As Boolean
    IsChoiceControl = ValueInList(controlValue, "choice_manual", "choice_formula")
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = TABLE_TYPE_BIVARIATE
End Function

Private Function ITableTypePolicy_HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasPercent = ValueInList(context.ColumnValue("percentage"), "row", "column", "total")
End Function

Private Function ITableTypePolicy_HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasTotal = True
End Function

Private Function ITableTypePolicy_HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasMissing = ValueInList(context.ColumnValue("missing"), "row", "column", "all")
End Function

Private Function ITableTypePolicy_HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasGraph = ValueInList(context.ColumnValue("graph"), "percentage", "values", "both")
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    Dim rowVariable As String
    Dim columnVariable As String

    rowVariable = context.RowVariable
    columnVariable = context.ColumnVariable

    If Not context.VariablesContains(rowVariable) Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    If Not context.VariablesContains(columnVariable) Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    If Not IsChoiceControl(context.VariableControl(rowVariable)) Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    If Not IsChoiceControl(context.VariableControl(columnVariable)) Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    ITableTypePolicy_IsValid = True
End Function

