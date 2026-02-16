VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressDisplayStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Lightweight progress display stub exposing Caption and Value properties for testing")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

Private mCaption As String
Private mValue As String

'@property-title Caption text surfaced by SetupImportService
Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal value As String)
    mCaption = value
End Property

'@property-title Value text surfaced by SetupImportService
Public Property Get Value() As String
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As String)
    mValue = newValue
End Property
