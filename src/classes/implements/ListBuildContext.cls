VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBuildContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("Concrete IListBuildContext capturing the worksheet and table metadata for a build")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements IListBuildContext

Private Type TContextState
    worksheetName As String
    tableTypeCode As Byte
    linelist As Object
End Type

Private this As TContextState
Private Const CLASSNAME As String = "ListBuildContext"

'@section Factory
'===============================================================================
'Create a context configured with the supplied worksheet name and table type.
Public Function Create(ByVal worksheetName As String, _
                       ByVal tableTypeCode As Byte, _
                       Optional ByVal linelist As Object = Nothing) As IListBuildContext

    ValidateWorksheetName worksheetName

    With New ListBuildContext
        .Initialise worksheetName, tableTypeCode, linelist
        Set Create = .Self
    End With
End Function

'Expose the current instance as its interface type.
Public Property Get Self() As IListBuildContext
    Set Self = Me
End Property

'Initialise the context; helper used by the factory.
Friend Sub Initialise(ByVal worksheetName As String, _
                      ByVal tableTypeCode As Byte, _
                      ByVal linelist As Object)

    ValidateWorksheetName worksheetName
    this.worksheetName = Trim$(worksheetName)
    this.tableTypeCode = tableTypeCode
    Set this.linelist = linelist
End Sub

'@section Validation helpers
'===============================================================================
Private Sub ValidateWorksheetName(ByVal worksheetName As String)
    If Trim$(worksheetName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "Worksheet name cannot be empty when creating a build context"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Property Get IListBuildContext_WorksheetName() As String
    IListBuildContext_WorksheetName = this.worksheetName
End Property

Private Property Get IListBuildContext_TableTypeCode() As Byte
    IListBuildContext_TableTypeCode = this.tableTypeCode
End Property

Private Property Get IListBuildContext_Linelist() As Object
    Set IListBuildContext_Linelist = this.linelist
End Property

