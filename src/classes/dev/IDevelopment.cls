VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing the development workflow coordinator"

'@Interface
'@Folder("Dev")
'@ModuleDescription("Interface describing the development workflow coordinator")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

'@class IDevelopment
'@description
'Defines the public contract for the Development class, exposing methods for
'managing code tables (classes, modules, forms, tests), bulk importing and
'exporting VBA components between disk and the host workbook, configuring
'deployment settings (hidden and protected sheets), and executing the full
'deployment workflow. Consumers obtain an IDevelopment reference from
'Development.Create and interact exclusively through this interface.
'@depends IPasswords


'@section Lifecycle
'===============================================================================
'@description Bind the development helper to the Dev worksheet and optional
'code sheet.

'@jump:initialise
'@sub-title Bind the development helper to a Dev worksheet and optional code sheet.
'@details
'Validates the devSheet, stores internal references, initialises collections,
'verifies required named ranges, and restores any previously registered code
'worksheet. When a codeSheet is provided, it replaces the restored worksheet.
'@param devSheet Worksheet. The worksheet hosting development named ranges.
'@param codeSheet Optional Worksheet. A worksheet to register as the code
'table host.
Public Sub Initialise(ByVal devSheet As Worksheet, Optional ByVal codeSheet As Worksheet)
End Sub


'@section Configuration
'===============================================================================
'@description Properties that expose and control the internal state of the
'development helper.

'@jump:dev-worksheet
'@prop-title Development worksheet reference.
'@details
'Returns the Dev worksheet that stores configuration named ranges such as
'folder paths and counter values. Set during Initialise.
'@return Worksheet. The development worksheet bound to this instance.
Public Property Get DevWorksheet() As Worksheet
End Property

'@jump:code-worksheet
'@prop-title Worksheet registered to host development tables.
'@details
'Returns the worksheet currently registered to host development ListObjects.
'May be Nothing when no code worksheet has been registered, in which case
'new tables are created on the Dev sheet itself.
'@return Worksheet. The code worksheet, or Nothing when none is registered.
Public Property Get CodeWorksheet() As Worksheet
End Property

'@jump:display-prompts
'@prop-title Whether confirmation prompts are displayed.
'@details
'Controls whether confirmation dialogs are shown before mutating operations
'such as importing, exporting, or adding tables. Defaults to True after
'Initialise. Set to False to suppress prompts during automated workflows.
'@return Boolean. True when prompts are enabled.
Public Property Get DisplayPrompts() As Boolean
End Property

'@jump:display-prompts-set
'@prop-title Set the confirmation prompt toggle.
'@param value Boolean. True to enable prompts, False to suppress them.
Public Property Let DisplayPrompts(ByVal value As Boolean)
End Property


'@section Worksheet Preparation
'===============================================================================
'@description Methods for registering code worksheets and creating
'development tables.

'@jump:add-code-sheets
'@sub-title Register an existing worksheet to host future development tables.
'@details
'Validates that the worksheet exists in the host workbook, shows a
'confirmation dialog when prompts are enabled, and persists the
'registration so it survives workbook close/reopen cycles.
'@param sheetName String. Name of the worksheet to register as the code
'table host.
'@return Worksheet. The registered worksheet reference, or Nothing when the
'user declines the confirmation prompt.
Public Function AddCodeSheets(ByVal sheetName As String) As Worksheet
End Function

'@jump:add-class-table
'@sub-title Create a classes table on the code worksheet.
'@details
'Creates a new ListObject with columns for class names and an interface
'Yes/No flag. The includeTests parameter switches the table tag between
'general classes and test classes, controlling the destination folder
'during code transfers.
'@param includeTests Optional Boolean. When True, creates a test-scoped
'classes table. Defaults to False.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
Public Function AddClassTable(Optional ByVal includeTests As Boolean = False) As ListObject
End Function

'@jump:add-module-table
'@sub-title Create a modules table on the code worksheet.
'@details
'Creates a new ListObject with a single column for module names. The
'includeTests parameter switches the tag between general modules and test
'modules, controlling the destination folder during transfers.
'@param includeTests Optional Boolean. When True, creates a test-scoped
'modules table. Defaults to False.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
Public Function AddModuleTable(Optional ByVal includeTests As Boolean = False) As ListObject
End Function

'@jump:add-forms-table
'@sub-title Create a forms mapping table on the code worksheet.
'@details
'Creates a new ListObject with two columns: "modules" and "corresponding
'form". Entries drive the AddFormsCodes and Deploy workflows, which copy
'code from a standard module into the code-behind of a UserForm component.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
Public Function AddFormsTable() As ListObject
End Function

'@jump:add-test-table
'@sub-title Create a test modules table on the code worksheet.
'@details
'Creates a new ListObject configured for test module entries with a single
'"modules" column. The table is tagged so that code transfers target the
'tests folder.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
Public Function AddTestTable() As ListObject
End Function


'@section Deployment
'===============================================================================
'@description Methods for configuring deployment settings and executing
'the deployment workflow.

'@jump:add-hidden-sheet
'@sub-title Register a worksheet that should be hidden during deployment.
'@details
'Appends the worksheet name to the hidden list. During Deploy, every entry
'is set to xlSheetVeryHidden.
'@param sheetName String. Name of the worksheet to hide during deployment.
Public Sub AddHiddenSheet(ByVal sheetName As String)
End Sub

'@jump:add-protected-sheet
'@sub-title Register a worksheet that should be protected during deployment.
'@details
'Adds the worksheet with its permission flags to the protected list.
'During Deploy, each entry is protected via the IPasswords handler with
'the recorded allowShapes and allowDeletingRows settings.
'@param sheetName String. Name of the worksheet to protect during deployment.
'@param allowShapes Optional Boolean. When True, shapes remain editable on
'the protected sheet. Defaults to True.
'@param allowDeletingRows Optional Boolean. When True, row deletion is
'allowed on the protected sheet. Defaults to True.
Public Sub AddProtectedSheet(ByVal sheetName As String, _
                             Optional ByVal allowShapes As Boolean = True, _
                             Optional ByVal allowDeletingRows As Boolean = True)
End Sub

'@jump:import-all
'@sub-title Import all code artefacts from configured folders into the host workbook.
'@details
'Iterates every development ListObject on host worksheets. For each table
'entry, the corresponding .bas or .cls file is imported from disk into the
'VBProject, replacing any existing component of the same name.
Public Sub ImportAll()
End Sub

'@jump:export-all
'@sub-title Export all code artefacts from the host workbook to configured folders.
'@details
'Iterates every development ListObject on host worksheets. For each table
'entry, the VBA component is exported from the VBProject to disk,
'overwriting any existing file.
Public Sub ExportAll()
End Sub

'@jump:add-forms-codes
'@sub-title Copy code modules into associated UserForm components.
'@details
'Synchronises form modules by copying the code from each source module
'listed in forms tables into the corresponding UserForm component. After
'copying, the source module is removed from the VBProject.
Public Sub AddFormsCodes()
End Sub

'@jump:deploy
'@sub-title Protect and hide registered worksheets, copy forms code, and set the deployment marker.
'@details
'Executes the full deployment workflow: applies the default password,
'synchronises form modules, hides registered hidden sheets, protects
'registered protected sheets using the IPasswords handler, optionally
'locks the workbook, hides the code worksheet, installs the debug exit
'handler, and sets the workbook-level deployment flag to "Yes".
'@param passwords IPasswords. The password handler used to apply protection.
'@param lockWorkbook Optional Boolean. When True, locks the host workbook
'structure after protecting sheets. Defaults to False.
Public Sub Deploy(ByVal passwords As IPasswords, Optional ByVal lockWorkbook As Boolean = False)
End Sub


'@section Status
'===============================================================================
'@description Read-only status query for the deployment state.

'@jump:in-deployment
'@sub-title Check whether the workbook has already been deployed.
'@details
'Reads the workbook-level named range "inDeployment" and returns True when
'the stored value equals "Yes" (case-insensitive). Returns False when the
'name does not exist or contains any other value.
'@return Boolean. True when the deployment flag equals "Yes".
Public Function InDeployment() As Boolean
End Function
