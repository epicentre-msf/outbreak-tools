VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsVariablesAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Wraps ILLVariables so policies can query controls and types")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsVariableAccessor

Private Type TVariableAdapterState
    variables As ILLVariables
End Type

Private this As TVariableAdapterState
Private Const CLASSNAME As String = "TableSpecsVariablesAdapter"

'@section Factory
'===============================================================================
'Create a new adapter around the provided variables reference.
Public Function Create(ByVal variables As ILLVariables) As ITableSpecsVariableAccessor
    ValidateVariables variables

    With New TableSpecsVariablesAdapter
        Set .Variables = variables
        Set Create = .Self
    End With
End Function

'Expose the current instance as the interface type.
Public Property Get Self() As ITableSpecsVariableAccessor
    Set Self = Me
End Property

'@section Properties
'===============================================================================
'Internal setter used by the factory once the reference is validated.
Public Property Set Variables(ByVal variables As ILLVariables)
    ValidateVariables variables
    Set this.variables = variables
End Property

'Access the wrapped variables reference.
Private Property Get Variables() As ILLVariables
    Set Variables = this.variables
End Property

'@section Validation helpers
'===============================================================================
'Ensure the adapter receives a usable variables reference.
Private Sub ValidateVariables(ByVal variables As ILLVariables)
    If variables Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "variables cannot be Nothing"
    End If

    If NormalizeValue(TypeName(variables)) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                   "variables must expose a concrete implementation"
    End If
End Sub

'Centralise error reporting logic.
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITableSpecsVariableAccessor_Contains(ByVal variableName As String) As Boolean
    ITableSpecsVariableAccessor_Contains = Variables.Contains(Trim$(variableName))
End Function

Private Function ITableSpecsVariableAccessor_Control(ByVal variableName As String) As String
    Dim valueText As String

    valueText = CStr(Variables.Value(colName:="control", varName:=Trim$(variableName)))
    ITableSpecsVariableAccessor_Control = Trim$(valueText)
End Function

Private Function ITableSpecsVariableAccessor_VariableType(ByVal variableName As String) As String
    Dim valueText As String

    valueText = CStr(Variables.Value(colName:="variable type", varName:=Trim$(variableName)))
    ITableSpecsVariableAccessor_VariableType = Trim$(valueText)
End Function

