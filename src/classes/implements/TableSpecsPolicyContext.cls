VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsPolicyContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Aggregates TableSpecs data needed by table-type policies")
'@depends ProjectError, TableSpecsPolicyHelpers
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsPolicyContext

Private Type TContextState
    tableTypeCode As Byte
    columnMap As ITableSpecsColumnMap
    variables As ITableSpecsVariableAccessor
    dictionary As ITableSpecsDictionaryAccessor
    rowVariable As String
    columnVariable As String
End Type

Private this As TContextState
Private Const CLASSNAME As String = "TableSpecsPolicyContext"

'@section Factory
'===============================================================================
'Create and configure a new context instance.
Public Function Create(ByVal tableTypeCode As Byte, _
                       ByVal columnMap As ITableSpecsColumnMap, _
                       ByVal variables As ITableSpecsVariableAccessor, _
                       ByVal dictionary As ITableSpecsDictionaryAccessor) As ITableSpecsPolicyContext

    With New TableSpecsPolicyContext
        .InitializeContext tableTypeCode, columnMap, variables, dictionary
        Set Create = .Self
    End With
End Function

'Return the instance as its interface type.
Public Property Get Self() As ITableSpecsPolicyContext
    Set Self = Me
End Property

'@section Initialisation
'===============================================================================
'Validate dependencies and store them locally for future queries.
Public Sub InitializeContext(ByVal tableTypeCode As Byte, _
                             ByVal columnMap As ITableSpecsColumnMap, _
                             ByVal variables As ITableSpecsVariableAccessor, _
                             ByVal dictionary As ITableSpecsDictionaryAccessor)

    ValidateColumnMap columnMap
    ValidateVariables variables
    ValidateDictionary dictionary

    this.tableTypeCode = tableTypeCode
    Set this.columnMap = columnMap
    Set this.variables = variables
    Set this.dictionary = dictionary

    CacheCoreVariables
End Sub

'Cache frequently requested values such as row/column variable identifiers.
Private Sub CacheCoreVariables()
    this.rowVariable = Trim$(ColumnValue("row"))
    this.columnVariable = Trim$(ColumnValue("column"))
End Sub

'@section Validation helpers
'===============================================================================
'Ensure the context receives a valid column map reference.
Private Sub ValidateColumnMap(ByVal columnMap As ITableSpecsColumnMap)
    If columnMap Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "columnMap cannot be Nothing"
    End If

    If NormalizeValue(TypeName(columnMap)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "columnMap must expose a concrete implementation"
    End If
End Sub

'Ensure the context receives a valid variables accessor.
Private Sub ValidateVariables(ByVal variables As ITableSpecsVariableAccessor)
    If variables Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "variables accessor cannot be Nothing"
    End If

    If NormalizeValue(TypeName(variables)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "variables accessor must expose a concrete implementation"
    End If
End Sub

'Ensure the context receives a valid dictionary accessor.
Private Sub ValidateDictionary(ByVal dictionary As ITableSpecsDictionaryAccessor)
    If dictionary Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "dictionary accessor cannot be Nothing"
    End If

    If NormalizeValue(TypeName(dictionary)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "dictionary accessor must expose a concrete implementation"
    End If
End Sub

'@section Helpers
'===============================================================================
'Read a value from the underlying column map while safely handling errors.
Private Function SafeColumnValue(ByVal columnName As String) As String
    Dim valueText As String

    On Error Resume Next
        valueText = CStr(this.columnMap.Value(columnName))
    On Error GoTo 0

    SafeColumnValue = Trim$(valueText)
End Function

'Consistent error helper.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsPolicyContext_TableTypeCode() As Byte
    ITableSpecsPolicyContext_TableTypeCode = this.tableTypeCode
End Function

Private Function ITableSpecsPolicyContext_ColumnValue(ByVal columnName As String) As String
    ITableSpecsPolicyContext_ColumnValue = SafeColumnValue(columnName)
End Function

Private Function ITableSpecsPolicyContext_RowVariable() As String
    ITableSpecsPolicyContext_RowVariable = this.rowVariable
End Function

Private Function ITableSpecsPolicyContext_ColumnVariable() As String
    ITableSpecsPolicyContext_ColumnVariable = this.columnVariable
End Function

Private Function ITableSpecsPolicyContext_VariablesContains(ByVal variableName As String) As Boolean
    ITableSpecsPolicyContext_VariablesContains = this.variables.Contains(variableName)
End Function

Private Function ITableSpecsPolicyContext_VariableControl(ByVal variableName As String) As String
    ITableSpecsPolicyContext_VariableControl = this.variables.Control(variableName)
End Function

Private Function ITableSpecsPolicyContext_VariableType(ByVal variableName As String) As String
    ITableSpecsPolicyContext_VariableType = this.variables.VariableType(variableName)
End Function

Private Function ITableSpecsPolicyContext_DictionaryVariableExists(ByVal variableName As String) As Boolean
    ITableSpecsPolicyContext_DictionaryVariableExists = this.dictionary.VariableExists(variableName)
End Function

