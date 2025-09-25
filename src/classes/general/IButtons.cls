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

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@Folder("Linelist")
'@ModuleDescription("Interface of Linelist Buttons")
'@Interface
'@label:IButtons
'@section Overview
'@details Declares the operations required to instantiate and format linelist buttons, and enumerates available button scopes.
'@jump:IButtons

Option Explicit

'@section Enumerations
'===============================================================================

'@details Describes the available button footprints that influence placement and formatting.
Public Enum ButtonScope
    ButtonScopeLarge = 1 'Large admin button footprint
    ButtonScopeSmall = 2 'Compact linelist button
    ButtonScopeGeo = 3   'Geographical linelist button
End Enum

'@section Interface
'===============================================================================

'@label:Add
'@jump:Add
'@sub-title Create a button shape at the configured location
'@details Implementations should create the shape only when absent and optionally wire a macro command and label.
'@param actionCommand Optional String macro to execute when the button is clicked.
'@param shapeLabel Optional String label to display atop the button.
Public Sub Add(Optional ByVal actionCommand As String = vbNullString, Optional ByVal shapeLabel As String = vbNullString): End Sub

'@label:Format
'@jump:Format
'@sub-title Apply a design provider to an existing button shape
'@details Delegates to an `ILLFormat` implementation so callers can maintain visual consistency.
'@param design ILLFormat formatter responsible for styling the shape.
'@depends ILLFormat
Public Sub Format(design As ILLFormat): End Sub

'@section Checkings
'===============================================================================

'@label:HasCheckings
'@jump:HasCheckings
'@sub-title Determine whether the button has recorded diagnostic entries
'@return Boolean True when at least one checking was logged.
'@depends Checking
Public Property Get HasCheckings() As Boolean: End Property

'@label:CheckingValues
'@jump:CheckingValues
'@sub-title Retrieve the recorded diagnostic entries
'@details Returns the underlying `IChecking` instance so callers can inspect messages.
'@return Object IChecking implementation containing the log.
'@depends Checking
Public Property Get CheckingValues() As Object: End Property
