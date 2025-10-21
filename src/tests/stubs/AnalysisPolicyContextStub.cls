VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisPolicyContextStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("ITableSpecsPolicyContext double exposing simple stubbed metadata")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITableSpecsPolicyContext

Private Type TContextStubState
    tableTypeCode As Byte
End Type

Private this As TContextStubState

Public Sub Configure(ByVal tableTypeCode As Byte)
    this.tableTypeCode = tableTypeCode
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsPolicyContext_TableTypeCode() As Byte
    ITableSpecsPolicyContext_TableTypeCode = this.tableTypeCode
End Function

Private Function ITableSpecsPolicyContext_ColumnValue(ByVal columnName As String) As String
    ITableSpecsPolicyContext_ColumnValue = vbNullString
End Function

Private Function ITableSpecsPolicyContext_RowVariable() As String
    ITableSpecsPolicyContext_RowVariable = vbNullString
End Function

Private Function ITableSpecsPolicyContext_ColumnVariable() As String
    ITableSpecsPolicyContext_ColumnVariable = vbNullString
End Function

Private Function ITableSpecsPolicyContext_VariablesContains(ByVal variableName As String) As Boolean
    ITableSpecsPolicyContext_VariablesContains = True
End Function

Private Function ITableSpecsPolicyContext_VariableControl(ByVal variableName As String) As String
    ITableSpecsPolicyContext_VariableControl = vbNullString
End Function

Private Function ITableSpecsPolicyContext_VariableType(ByVal variableName As String) As String
    ITableSpecsPolicyContext_VariableType = vbNullString
End Function

Private Function ITableSpecsPolicyContext_DictionaryVariableExists(ByVal variableName As String) As Boolean
    ITableSpecsPolicyContext_DictionaryVariableExists = True
End Function
