VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CloneableItemStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Test stub exposing a Clone method for array utility tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Type TStubState
    identifier As String
End Type

Private this As TStubState

'@section Configuration
'===============================================================================
Public Sub Configure(ByVal identifier As String)
    this.identifier = identifier
End Sub

'@section Accessors
'===============================================================================
Public Property Get Identifier() As String
    Identifier = this.identifier
End Property

'@section Clone support
'===============================================================================
Public Function Clone() As CloneableItemStub
    Dim copy As CloneableItemStub
    Set copy = New CloneableItemStub
    copy.Configure this.identifier
    Set Clone = copy
End Function
