VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicyTimeSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Implements TableSpecs rules for time-series tables")
'@depends TableSpecsPolicyHelpers
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableTypePolicy

Private Const CLASSNAME As String = "TableTypePolicyTimeSeries"

'@section Helpers
'===============================================================================
'Determine whether the percentage selection is row-based or column-based.
Private Function HasRowOrColumnPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    HasRowOrColumnPercent = ValueInList(context.ColumnValue("percentage"), "row", "column")
End Function

'Validate whether the supplied control is compatible with analysis tables.
Private Function IsChoiceControl(ByVal controlValue As String) As Boolean
    IsChoiceControl = ValueInList(controlValue, "choice_manual", "choice_formula")
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = TABLE_TYPE_TIME_SERIES
End Function

Private Function ITableTypePolicy_HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    If Not HasRowOrColumnPercent(context) Then
        ITableTypePolicy_HasPercent = False
        Exit Function
    End If

    ITableTypePolicy_HasPercent = ITableTypePolicy_HasTotal(context)
End Function

Private Function ITableTypePolicy_HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
    Dim hasExplicitTotal As Boolean
    Dim columnVariable As String

    columnVariable = context.ColumnValue("column")
    hasExplicitTotal = ValueEquals(context.ColumnValue("total"), "yes")

    ITableTypePolicy_HasTotal = (hasExplicitTotal Or HasRowOrColumnPercent(context)) And _
                                HasText(columnVariable)
End Function

Private Function ITableTypePolicy_HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasMissing = ValueEquals(context.ColumnValue("missing"), "yes") And _
                                  HasText(context.ColumnValue("column"))
End Function

Private Function ITableTypePolicy_HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasGraph = False
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    Dim rowVariable As String
    Dim columnVariable As String
    Dim rowType As String

    rowVariable = context.RowVariable
    columnVariable = context.ColumnVariable
    rowType = context.VariableType(rowVariable)

    If Not ValueEquals(rowType, "date") Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    If HasText(columnVariable) Then
        ITableTypePolicy_IsValid = IsChoiceControl(context.VariableControl(columnVariable))
    Else
        ITableTypePolicy_IsValid = True
    End If
End Function

