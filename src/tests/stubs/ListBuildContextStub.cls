VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBuildContextStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of IListBuildContext for coordinator tests")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements IListBuildContext

Private Type TContextState
    worksheetName As String
    tableTypeCode As Byte
    linelist As Object
End Type

Private this As TContextState

Public Sub Configure(ByVal worksheetName As String, ByVal tableTypeCode As Byte)
    this.worksheetName = worksheetName
    this.tableTypeCode = tableTypeCode
End Sub

Public Property Set Linelist(ByVal linelist As Object)
    Set this.linelist = linelist
End Property

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

