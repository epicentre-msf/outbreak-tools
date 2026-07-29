VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressDisplayStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Progress display stub exposing Caption and Value for the setup import tests"

'@Folder("Tests.Helpers")
'@ModuleDescription("Progress display stub exposing Caption and Value for the setup import tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@class ProgressDisplayStub
'@description
'Stands in for the progress label of the Imports form. SetupImportService asks
'a display object for a Caption or a Value before it writes to it, and this
'stub answers both, so a test can read back whatever the service wrote.

'The two fields carry a name of their own. A field spelled like the property
'that reads it makes the VBA editor renormalise one to the other, and that is a
'known way to freeze it.
Private Type TProgressDisplay
    displayCaption As String
    displayValue As String
End Type

Private this As TProgressDisplay

'@label:caption
'@prop-title Text the service last wrote to the caption.
'@return String. The stored caption.
Public Property Get Caption() As String
    Caption = this.displayCaption
End Property

'@label:caption-set
'@prop-title Store the caption text.
'@param newCaption String. The text to store.
Public Property Let Caption(ByVal newCaption As String)
    this.displayCaption = newCaption
End Property

'@label:value
'@prop-title Text the service last wrote to the value.
'@return String. The stored value.
Public Property Get Value() As String
    Value = this.displayValue
End Property

'@label:value-set
'@prop-title Store the value text.
'@param newValue String. The text to store.
Public Property Let Value(ByVal newValue As String)
    this.displayValue = newValue
End Property
