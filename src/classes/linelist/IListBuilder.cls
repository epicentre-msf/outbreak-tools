VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IListBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Contract for building HList or VList data entry sheets"

'@Interface
'@Folder("Linelist")
'@ModuleDescription("Contract for building HList or VList data entry sheets")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@class IListBuilder
'@description
'Defines the public contract for the ListBuilder class, exposing a single
'Build method that constructs an HList (horizontal) or VList (vertical) data
'entry sheet with formatting, freeze panes, sections, buttons, ListObjects,
'and worksheet protection.

'ListBuilderLayer modes: HList is horizontal, VList is vertical.
Public Enum ListBuilderLayer
    ListBuilderLayerHList = 1
    ListBuilderLayerVList = 2
End Enum


'@section Operations
'===============================================================================
'@description Data entry sheet construction.

'@jump:build
'@sub-title Build the data entry sheet
'@details
'Dispatches to the appropriate builder based on the layer mode set during
'factory construction. HList mode creates a horizontal data entry sheet with
'a ListObject table, printed companion sheet, filtered companion sheet, buttons,
'and pivot table. VList mode creates a vertical data entry sheet with sections,
'a GoTo dropdown, and PLAGEVALUES range reference. Both modes apply formatting,
'freeze panes, build sections via SectionBuilder, and protect the sheet.
Public Sub Build()
End Sub


'@section Checkings
'===============================================================================
'@description Diagnostic logging access for sheet build results.

'@jump:has-checkings
'@prop-title Whether the builder has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean
End Property

'@jump:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return IChecking. The checking instance, or Nothing.
Public Property Get CheckingValues() As IChecking
End Property
