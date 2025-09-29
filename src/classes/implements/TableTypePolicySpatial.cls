VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableTypePolicySpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Implements TableSpecs rules for spatial tables")
'@depends TableSpecsPolicyHelpers
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableTypePolicy

Private Const CLASSNAME As String = "TableTypePolicySpatial"

'@section Helpers
'===============================================================================
'Determine whether the supplied control is supported by the analysis rules.
Private Function IsChoiceControl(ByVal controlValue As String) As Boolean
    IsChoiceControl = ValueInList(controlValue, "choice_manual", "choice_formula")
End Function

'Check whether the dictionary exposes any of the expected spatial prefixes.
Private Function HasSpatialVariable(ByVal context As ITableSpecsPolicyContext, ByVal baseName As String) As Boolean
    HasSpatialVariable = context.DictionaryVariableExists("adm1_" & baseName) Or _
                         context.DictionaryVariableExists("hf_" & baseName)
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableTypePolicy_TableTypeCode() As Byte
    ITableTypePolicy_TableTypeCode = TABLE_TYPE_SPATIAL
End Function

Private Function ITableTypePolicy_HasPercent(ByVal context As ITableSpecsPolicyContext) As Boolean
    If Not ValueEquals(context.ColumnValue("percentage"), "yes") Then
        ITableTypePolicy_HasPercent = False
        Exit Function
    End If

    ITableTypePolicy_HasPercent = ITableTypePolicy_HasTotal(context)
End Function

Private Function ITableTypePolicy_HasTotal(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasTotal = HasText(context.ColumnValue("column"))
End Function

Private Function ITableTypePolicy_HasMissing(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasMissing = ValueEquals(context.ColumnValue("missing"), "yes") And _
                                  HasText(context.ColumnValue("column"))
End Function

Private Function ITableTypePolicy_HasGraph(ByVal context As ITableSpecsPolicyContext) As Boolean
    ITableTypePolicy_HasGraph = ValueEquals(context.ColumnValue("graph"), "yes")
End Function

Private Function ITableTypePolicy_IsValid(ByVal context As ITableSpecsPolicyContext) As Boolean
    Dim rowVariable As String
    Dim columnVariable As String

    rowVariable = context.RowVariable
    columnVariable = context.ColumnVariable

    If Not HasSpatialVariable(context, rowVariable) Then
        ITableTypePolicy_IsValid = False
        Exit Function
    End If

    If HasText(columnVariable) Then
        ITableTypePolicy_IsValid = IsChoiceControl(context.VariableControl(columnVariable))
    Else
        ITableTypePolicy_IsValid = True
    End If
End Function

