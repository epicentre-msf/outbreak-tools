VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FakeTableSpecsPolicyContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests")
'@ModuleDescription("Test double implementing ITableSpecsPolicyContext for policy unit tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsPolicyContext

Private Type TContextData
    tableTypeCode As Byte
    values As Collection
    rowVariable As String
    columnVariable As String
    presentVariables As Collection
    controls As Collection
    variableTypes As Collection
    dictionaryVariables As Collection
End Type

Private this As TContextData

'@section Lifecycle
'===============================================================================
Private Sub Class_Initialize()
    Set this.values = New Collection
    Set this.presentVariables = New Collection
    Set this.controls = New Collection
    Set this.variableTypes = New Collection
    Set this.dictionaryVariables = New Collection
End Sub

'@section Configuration helpers
'===============================================================================
'Apply the basic configuration shared across tests.
Public Sub Configure(ByVal tableTypeCode As Byte, _
                     ByVal rowVariable As String, _
                     ByVal columnVariable As String)
    this.tableTypeCode = tableTypeCode
    this.rowVariable = rowVariable
    this.columnVariable = columnVariable
End Sub

'Store or replace a table column value.
Public Sub SetColumnValue(ByVal columnName As String, ByVal valueText As String)
    SetCollectionValue this.values, NormalizeValue(columnName), valueText
End Sub

'Mark a variable as present (or absent) inside the fake variable accessor.
Public Sub SetVariablePresence(ByVal variableName As String, ByVal isPresent As Boolean)
    Dim key As String
    key = NormalizeValue(variableName)

    If isPresent Then
        EnsureCollectionItem this.presentVariables, key, CStr(True)
    Else
        RemoveCollectionItem this.presentVariables, key
    End If
End Sub

'Store the control associated to a variable.
Public Sub SetVariableControl(ByVal variableName As String, ByVal controlValue As String)
    SetCollectionValue this.controls, NormalizeValue(variableName), controlValue
End Sub

'Store the variable type associated to a variable.
Public Sub SetVariableType(ByVal variableName As String, ByVal typeValue As String)
    SetCollectionValue this.variableTypes, NormalizeValue(variableName), typeValue
End Sub

'Expose a dictionary variable so spatial policies can validate prefixes.
Public Sub AddDictionaryVariable(ByVal variableName As String)
    EnsureCollectionItem this.dictionaryVariables, NormalizeValue(variableName), variableName
End Sub

'@section Internal helpers
'===============================================================================
'Insert or update a key/value pair inside a collection.
Private Sub SetCollectionValue(ByRef target As Collection, ByVal key As String, ByVal valueText As String)
    RemoveCollectionItem target, key
    On Error Resume Next
        target.Add valueText, key
    On Error GoTo 0
End Sub

'Safely remove a keyed entry when it exists.
Private Sub RemoveCollectionItem(ByRef target As Collection, ByVal key As String)
    On Error Resume Next
        target.Remove key
    On Error GoTo 0
End Sub

'Ensure a value exists inside a collection keyed by the provided identifier.
Private Sub EnsureCollectionItem(ByRef target As Collection, ByVal key As String, ByVal valueText As String)
    On Error Resume Next
        target.Add valueText, key
    On Error GoTo 0
End Sub

'Retrieve a stored value or return an empty string.
Private Function CollectionValue(ByVal source As Collection, ByVal key As String) As String
    Dim valueText As String

    On Error Resume Next
        valueText = CStr(source.Item(key))
    On Error GoTo 0

    CollectionValue = valueText
End Function

'Check whether a keyed value exists inside the collection.
Private Function CollectionContains(ByVal source As Collection, ByVal key As String) As Boolean
    Dim dummy As Variant

    On Error Resume Next
        dummy = source.Item(key)
    CollectionContains = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsPolicyContext_TableTypeCode() As Byte
    ITableSpecsPolicyContext_TableTypeCode = this.tableTypeCode
End Function

Private Function ITableSpecsPolicyContext_ColumnValue(ByVal columnName As String) As String
    ITableSpecsPolicyContext_ColumnValue = CollectionValue(this.values, NormalizeValue(columnName))
End Function

Private Function ITableSpecsPolicyContext_RowVariable() As String
    ITableSpecsPolicyContext_RowVariable = this.rowVariable
End Function

Private Function ITableSpecsPolicyContext_ColumnVariable() As String
    ITableSpecsPolicyContext_ColumnVariable = this.columnVariable
End Function

Private Function ITableSpecsPolicyContext_VariablesContains(ByVal variableName As String) As Boolean
    ITableSpecsPolicyContext_VariablesContains = CollectionContains(this.presentVariables, NormalizeValue(variableName))
End Function

Private Function ITableSpecsPolicyContext_VariableControl(ByVal variableName As String) As String
    ITableSpecsPolicyContext_VariableControl = CollectionValue(this.controls, NormalizeValue(variableName))
End Function

Private Function ITableSpecsPolicyContext_VariableType(ByVal variableName As String) As String
    ITableSpecsPolicyContext_VariableType = CollectionValue(this.variableTypes, NormalizeValue(variableName))
End Function

Private Function ITableSpecsPolicyContext_DictionaryVariableExists(ByVal variableName As String) As Boolean
    ITableSpecsPolicyContext_DictionaryVariableExists = CollectionContains(this.dictionaryVariables, NormalizeValue(variableName))
End Function

