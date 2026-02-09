VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Linelist Buttons"

'@Interface
'@Folder("General")
'@ModuleDescription("Interface of Linelist Buttons")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, ProcedureNotUsed, SuperfluousAnnotationArgument

Option Explicit

'@class IButtons
'@description
'Defines the public contract for the Buttons class, exposing
'worksheet shape creation, placement, and formatting for linelist
'buttons. Supports large, small, and geo button scopes.


'@section Enumerations
'===============================================================================
'@description Available button footprints.

Public Enum ButtonScope
    ButtonScopeLarge = 1
    ButtonScopeSmall = 2
    ButtonScopeGeo = 3
End Enum


'@section Properties
'===============================================================================
'@description State accessors.

'@jump:output-range
'@prop-title Anchor range used for placement
'@return Range. Top-left anchor cell.
Public Property Get OutputRange() As Range: End Property

'@jump:output-range-set
'@prop-title Update the anchor range used for placement
'@param rng Range. Providing the top-left anchor cell.
Public Property Set OutputRange(ByVal rng As Range): End Property

'@jump:name
'@prop-title Name of the shape in the worksheet
'@return String. Shape identifier.
Public Property Get Name() As String: End Property

'@jump:shape-scope
'@prop-title Formatting scope applied to the button
'@return Byte. ButtonScope value.
Public Property Get ShapeScope() As Byte: End Property


'@section Operations
'===============================================================================
'@description Button creation and formatting.

'@jump:add
'@sub-title Create a button shape at the configured location
'@param actionCommand String. Macro to execute on click. Defaults to vbNullString.
'@param shapeLabel String. Label displayed on the button. Defaults to vbNullString.
Public Sub Add(Optional ByVal actionCommand As String = vbNullString, _
               Optional ByVal shapeLabel As String = vbNullString): End Sub

'@jump:format
'@sub-title Apply a design provider to an existing button shape
'@param design ILLFormat. Formatter responsible for styling the shape.
Public Sub Format(design As ILLFormat): End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access.

'@jump:has-checkings
'@prop-title Whether diagnostic entries have been recorded
'@return Boolean. True when at least one entry exists.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. The checking instance, or Nothing.
Public Property Get CheckingValues() As Object: End Property
