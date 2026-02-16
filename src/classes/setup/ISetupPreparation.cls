VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface orchestrating setup preparation tasks"

'@Interface
'@Folder("Setup")
'@ModuleDescription("Describes the preparation workflow responsible for dropdown registration and update watcher initialisation")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class ISetupPreparation
'@description
'Defines the public contract for the SetupPreparation class, exposing
'methods to bind a preparation helper to a workbook, execute the full
'preparation workflow (dropdowns, updated values registry, validations,
'hidden/protected sheets), and access the resulting dropdown manager
'and update watcher objects.
'@depends IDevelopment, IDropdownLists, IUpdatedValues


'@section Lifecycle
'===============================================================================

'@jump:configure
'@sub-title Bind the preparation helper to a workbook hosting setup worksheets
'@details
'Stores the workbook reference and resets cached worksheet and manager
'objects. Must be called before any other method or property on this
'instance.
'@param hostBook Workbook. The workbook owning the setup assets.
'@export
Public Sub Configure(ByVal hostBook As Workbook)
End Sub

'@jump:prepare
'@sub-title Execute the full preparation workflow
'@details
'Runs the complete preparation sequence: registers all dropdown list
'objects, initialises the updated values registry, applies data
'validations to setup tables (Dictionary, Exports, Analysis), and
'ensures hidden and protected sheet configuration via the supplied
'development manager.
'@param manager IDevelopment. Development manager providing hidden/protected sheet registration.
'@export
Public Sub Prepare(ByRef manager As IDevelopment)
End Sub

'@section Operations
'===============================================================================

'@jump:ensure-dropdowns
'@sub-title Ensure dropdown list objects exist with the expected default entries
'@details
'Creates the dropdown manager if needed, then registers all standard
'dropdown lists (yes/no, formats, variable types, etc.) on the dropdown
'worksheet. Safe to call multiple times; subsequent calls are no-ops
'when the manager is already initialised.
'@export
Public Sub EnsureDropdowns()
End Sub

'@jump:ensure-updated-registry
'@sub-title Initialise the updated values registry and register watched tables
'@details
'Creates the update watcher if needed, clears any existing registry data,
'registers the watched setup sheets (Dictionary, Choices, Analysis, Exports),
'and resets all status tags to their default values.
'@export
Public Sub EnsureUpdatedRegistry()
End Sub

'@section State
'===============================================================================

'@jump:dropdowns
'@prop-title Dropdown manager created during preparation
'@details
'Returns the IDropdownLists manager bound to the dropdown worksheet.
'Lazily initialises the manager on first access if it was not already
'created by EnsureDropdowns or Prepare.
'@return IDropdownLists. The dropdown manager instance.
'@export
Public Property Get Dropdowns() As IDropdownLists
End Property

'@jump:updated-watcher
'@prop-title Updated values watcher created during preparation
'@details
'Returns the IUpdatedValues watcher bound to the registry worksheet.
'Lazily initialises the watcher on first access if it was not already
'created by EnsureUpdatedRegistry or Prepare.
'@return IUpdatedValues. The update watcher instance.
'@export
Public Property Get UpdatedWatcher() As IUpdatedValues
End Property

'@jump:host-workbook
'@prop-title Workbook hosting the setup worksheets
'@details
'Returns the workbook reference stored during Configure. Raises an
'error when the instance has not been configured.
'@return Workbook. The bound workbook.
'@export
Public Property Get HostWorkbook() As Workbook
End Property
