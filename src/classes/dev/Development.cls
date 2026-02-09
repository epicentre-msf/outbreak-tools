VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Development"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook development helper orchestrating code tables, code transfers, and deployment"

'@PredeclaredId
'@Folder("Dev")
'@ModuleDescription("Workbook development helper orchestrating code tables, code transfers, and deployment")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit
Implements IDevelopment

'@class Development
'@description
'The Development class is the central orchestrator for workbook development
'workflows in OutbreakTools. It manages code tables (ListObjects that list
'VBA classes, modules, forms, and tests), drives bulk import and export of
'code artefacts between disk and the host workbook VBProject, and executes
'deployment by protecting and hiding registered worksheets. Consumers
'interact through the IDevelopment interface using the predeclared factory
'pattern (Development.Create).
'@depends BetterArray, IPasswords
'@version 1.0 (2026-02-09)
'@author Yves Amevoin


Private Type TDevelopment
    hostwb As Workbook             'Host workbook containing development assets
    devSheet As Worksheet          'Development sheet carrying named ranges for code folders
    codeSheet As Worksheet         'Worksheet collecting classes/modules/forms/test tables
    hiddenSheets As BetterArray    'Sheets to hide during deployment
    protectedSheets As BetterArray 'Sheets to protect during deployment
    displayPrompts As Boolean      'Toggle for confirmation prompts
End Type

Private this As TDevelopment

Private Const CLASSNAME As String = "Development"
Private Const CONFIRM_TITLE As String = "Development helper"

Private Const MODULE_CODES_NAME As String = "ModulesCodes"
Private Const CLASS_CODES_NAME As String = "ClassesImplementation"
Private Const TEST_CODES_NAME As String = "TestsCodes"
Private Const CODE_SHEET_NAMEDEF As String = "Development_CodeSheet"
Private Const WORKBOOK_DEPLOYMENT_FLAG As String = "inDeployment"

Private Const COUNTER_CLASSES As String = "Development_ClassesCounter"
Private Const COUNTER_MODULES As String = "Development_ModulesCounter"
Private Const COUNTER_FORMS As String = "Development_FormsCounter"
Private Const COUNTER_TESTS As String = "Development_TestsCounter"

Private Const PREFIX_CLASSES As String = "ClassesLo"
Private Const PREFIX_MODULES As String = "ModulesLo"
Private Const PREFIX_FORMS As String = "FormsLo"
Private Const PREFIX_TESTS As String = "TestsLo"

Private Const TAG_CLASSES As String = "general classes"
Private Const TAG_MODULES As String = "general modules"
Private Const TAG_FORMS As String = "general form modules"
Private Const TAG_TESTS As String = "tests modules"
Private Const TAG_TEST_CLASSES As String = "tests classes"

Private Const YES_NO_LIST As String = "Yes,No"
Private Const INFO_RANGE As String = "Informations"
Private Const DEFAULT_ANCHOR As String = "E5"

Private Enum DevelopmentTransferScope
    transferModules = 1
    transferClasses = 2
End Enum

Private Enum DevelopmentTransferMode
    transferImport = 1
    transferExport = 2
End Enum

Private Type TFolderSettings
    modulesPath As String
    classesPath As String
    testsPath As String
End Type


'@section Factory
'===============================================================================
'@description Predeclared factory entry point for creating Development instances.

'@label:create
'@sub-title Create a Development instance.
'@details
'Factory method on the predeclared instance that allocates a new Development
'object and binds it to the supplied Dev worksheet via Initialise. When a
'codeSheet is provided, it is registered as the target for new development
'tables. The returned object is cast to the IDevelopment interface so
'consumers program against the contract rather than the concrete class.
'@param devSheet Worksheet. The worksheet hosting development named ranges
'such as module and class folder paths.
'@param codeSheet Optional Worksheet. A separate worksheet to host code
'ListObjects. When omitted, tables are placed on devSheet.
'@return IDevelopment. A fully initialised instance ready for use.
'@throws ProjectError.ObjectNotInitialized When devSheet is Nothing.
'@throws ProjectError.ElementNotFound When required named ranges are missing
'on the Dev worksheet.
'@depends BetterArray
'@export
Public Function Create(ByVal devSheet As Worksheet, Optional ByVal codeSheet As Worksheet) As IDevelopment
    Dim instance As Development

    Set instance = New Development
    instance.Initialise devSheet, codeSheet
    Set Create = instance
End Function


'@section Initialisation
'===============================================================================
'@description Instance setup that binds internal state to the host workbook.

'@label:initialise
'@sub-title Bind this instance to a Dev worksheet and optional code sheet.
'@details
'Validates that devSheet is not Nothing, stores the worksheet and its parent
'workbook, enables user confirmation prompts, and initialises the hidden and
'protected sheet collections. Required named ranges for module and class
'folder paths are verified on the Dev sheet. Any previously registered code
'worksheet is restored from a persisted worksheet-level name. When a
'codeSheet argument is supplied, it replaces the restored worksheet.
'@param devSheet Worksheet. The development worksheet carrying configuration
'named ranges.
'@param codeSheet Optional Worksheet. A worksheet to register as the code
'table host. When omitted, any previously persisted code worksheet is used.
'@throws ProjectError.ObjectNotInitialized When devSheet is Nothing.
'@throws ProjectError.ElementNotFound When required named ranges are missing.
Public Sub Initialise(ByVal devSheet As Worksheet, Optional ByVal codeSheet As Worksheet)

    If devSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Development worksheet reference is required"
    End If

    Set this.devSheet = devSheet
    Set this.hostwb = devSheet.Parent
    this.displayPrompts = True

    EnsureCollections
    EnsureNamedRanges

    'Resume using previously registered code worksheet when workbook is reopened.
    RestoreCodeWorksheet

    If Not codeSheet Is Nothing Then
        RegisterCodeWorksheet codeSheet.Name
    End If
End Sub


'@section Sheet Preparation API
'===============================================================================
'@description Public methods for registering worksheets and creating
'development tables (classes, modules, forms, tests).

'@label:add-code-sheets
'@sub-title Register an existing worksheet to host future development tables.
'@details
'Validates that the supplied name is non-empty and that the worksheet exists
'in the host workbook. When user prompts are enabled, a confirmation dialog
'is displayed before proceeding. The registered worksheet is persisted via a
'worksheet-level named range so it survives workbook close/reopen cycles.
'@param sheetName String. Name of the worksheet to register as the code
'table host.
'@return Worksheet. The registered worksheet reference, or Nothing when the
'user declines the confirmation prompt.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@throws ProjectError.InvalidArgument When sheetName is empty.
'@throws ProjectError.ElementNotFound When the worksheet does not exist in
'the host workbook.
'@Description("Register an existing worksheet so future tables are created on it.")
Public Function AddCodeSheets(ByVal sheetName As String) As Worksheet
    EnsureInitialised

    Dim trimmedName As String
    trimmedName = Trim$(sheetName)

    If LenB(trimmedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Code worksheet name cannot be empty"
    End If

    Dim sheet As Worksheet
    Set sheet = EnsureWorksheet(trimmedName, "code")

    If Not RequestConfirmation("register code worksheet '" & sheet.Name & "'") Then Exit Function

    RegisterCodeWorksheet sheet.Name
    Set AddCodeSheets = sheet
End Function

'@label:add-class-table
'@sub-title Create a classes table on the code worksheet.
'@details
'Creates a new ListObject with columns for class names and an interface
'Yes/No flag. The includeTests parameter switches the table tag between
'"general classes" and "tests classes", which determines the destination
'folder during code transfers. A confirmation dialog is shown when prompts
'are enabled. The table is placed at the next available horizontal anchor
'on the code worksheet (or the Dev sheet when no code worksheet is registered).
'@param includeTests Optional Boolean. When True, creates a test-scoped
'classes table. Defaults to False.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@Description("Create a classes table, optionally scoped to tests.")
Public Function AddClassTable(Optional ByVal includeTests As Boolean = False) As ListObject
    EnsureInitialised

    If Not RequestConfirmation("add classes table") Then Exit Function

    Dim tagValue As String
    tagValue = IIf(includeTests, TAG_TEST_CLASSES, TAG_CLASSES)

    Set AddClassTable = CreateListObject(Array("classes", "hasinterface"), _
                                         tagValue, _
                                         COUNTER_CLASSES, _
                                         PREFIX_CLASSES, _
                                         addYesNoValidation:=True)
End Function

'@label:add-module-table
'@sub-title Create a modules table on the code worksheet.
'@details
'Creates a new ListObject with a single column for module names. The
'includeTests parameter switches the table tag between "general modules"
'and "tests modules", controlling the destination folder during transfers.
'A confirmation dialog is shown when prompts are enabled.
'@param includeTests Optional Boolean. When True, creates a test-scoped
'modules table. Defaults to False.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@Description("Create a modules table, optionally scoped to tests.")
Public Function AddModuleTable(Optional ByVal includeTests As Boolean = False) As ListObject
    EnsureInitialised

    If Not RequestConfirmation("add modules table") Then Exit Function

    Dim tagValue As String
    tagValue = IIf(includeTests, TAG_TESTS, TAG_MODULES)

    Set AddModuleTable = CreateListObject(Array("modules"), _
                                          tagValue, _
                                          COUNTER_MODULES, _
                                          PREFIX_MODULES)
End Function

'@label:add-forms-table
'@sub-title Create a forms mapping table on the code worksheet.
'@details
'Creates a new ListObject with two columns: "modules" and "corresponding
'form". Entries in this table drive the AddFormsCodes and Deploy workflows,
'which copy code from a standard module into the code-behind of a UserForm
'component. A confirmation dialog is shown when prompts are enabled.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@Description("Create a forms mapping table used for code transfers to userforms.")
Public Function AddFormsTable() As ListObject
    EnsureInitialised

    If Not RequestConfirmation("add forms modules table") Then Exit Function

    Set AddFormsTable = CreateListObject(Array("modules", "corresponding form"), _
                                         TAG_FORMS, _
                                         COUNTER_FORMS, _
                                         PREFIX_FORMS)
End Function

'@label:add-test-table
'@sub-title Create a test modules table on the code worksheet.
'@details
'Creates a new ListObject configured for test module entries with a single
'"modules" column. The table is tagged as "tests modules" so that code
'transfers target the tests folder. A confirmation dialog is shown when
'prompts are enabled.
'@return ListObject. The newly created table, or Nothing when the user
'declines the confirmation prompt.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@Description("Create a test modules table for organising test imports/exports.")
Public Function AddTestTable() As ListObject
    EnsureInitialised

    If Not RequestConfirmation("add tests table") Then Exit Function

    Set AddTestTable = CreateListObject(Array("modules"), _
                                        TAG_TESTS, _
                                        COUNTER_TESTS, _
                                        PREFIX_TESTS)
End Function


'@section Deployment Configuration
'===============================================================================
'@description Methods for registering worksheets that must be hidden or
'protected when the workbook is deployed.

'@label:add-hidden-sheet
'@sub-title Track a worksheet that must be hidden during deployment.
'@details
'Appends the supplied worksheet name to the internal hiddenSheets
'collection. During Deploy, every entry in this collection is set to
'xlSheetVeryHidden. The worksheet must exist in the host workbook.
'@param sheetName String. Name of the worksheet to hide during deployment.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@throws ProjectError.InvalidArgument When sheetName is empty.
'@throws ProjectError.ElementNotFound When the worksheet does not exist.
'@Description("Track a worksheet that must be hidden during deployment.")
Public Sub AddHiddenSheet(ByVal sheetName As String)
    EnsureInitialised
    EnsureCollections

    Dim sheet As Worksheet
    Set sheet = EnsureWorksheet(sheetName, "hidden")

    this.hiddenSheets.Push sheet.Name
End Sub

'@label:add-protected-sheet
'@sub-title Track a worksheet that needs protection during deployment.
'@details
'Registers a worksheet by name so that it will be protected when Deploy is
'called. The optional parameters control whether shapes and row deletion
'are permitted on the protected sheet. The registration is stored as a
'BetterArray triplet (name, allowShapes, allowDeletingRows) inside the
'protectedSheets collection.
'@param sheetName String. Name of the worksheet to protect during deployment.
'@param allowShapes Optional Boolean. When True, shapes remain editable on
'the protected sheet. Defaults to True.
'@param allowDeletingRows Optional Boolean. When True, row deletion is
'allowed on the protected sheet. Defaults to True.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@throws ProjectError.InvalidArgument When sheetName is empty.
'@throws ProjectError.ElementNotFound When the worksheet does not exist.
'@Description("Track a worksheet that needs to be protected during deployment.")
Public Sub AddProtectedSheet(ByVal sheetName As String,  _
                          Optional ByVal allowShapes As Boolean = True, _
                          Optional ByVal allowDeletingRows As Boolean = True)
    EnsureInitialised
    EnsureCollections

    Dim sheet As Worksheet
    Set sheet = EnsureWorksheet(sheetName, "protected")

    RegisterSheet this.protectedSheets, sheet.Name, allowShapes, allowDeletingRows
End Sub


'@section Code Transfer
'===============================================================================
'@description Bulk import and export of VBA components listed in
'development tables.

'@label:import-all
'@sub-title Import every module and class listed in Dev tables into the host workbook.
'@details
'Iterates every ListObject on the Dev sheet and any registered code
'worksheet. Each table tag is read to determine the transfer scope (modules
'or classes) and the target disk folder. For each entry in the table, the
'corresponding .bas or .cls file is imported from disk into the host
'workbook VBProject, replacing any existing component of the same name. A
'timestamped status message is written to the Dev sheet Informations range
'upon completion. Requires that host sheets are unlocked.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@throws ProjectError.ErrorUnexpectedState When a host sheet is protected.
'@Description("Import every module/class listed in Dev tables into this workbook.")
Public Sub ImportAll()
    EnsureInitialised
    EnsureHostSheetsUnlocked

    If Not RequestConfirmation("import code artefacts into workbook '" & this.hostwb.Name & "'") Then Exit Sub

    ProcessAllTables transferImport
    WriteInformation "Finished Imports At: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
End Sub

'@label:export-all
'@sub-title Export every module and class listed in Dev tables to disk.
'@details
'Iterates every ListObject on the Dev sheet and any registered code
'worksheet. Each table tag is read to determine the transfer scope and the
'source disk folder. For each entry, the VBA component is exported from
'the host workbook VBProject to disk, overwriting any existing file. A
'timestamped status message is written upon completion. Requires that host
'sheets are unlocked.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@throws ProjectError.ErrorUnexpectedState When a host sheet is protected.
'@Description("Export every module/class listed in Dev tables to disk.")
Public Sub ExportAll()
    EnsureInitialised
    EnsureHostSheetsUnlocked

    If Not RequestConfirmation("export code artefacts from workbook '" & this.hostwb.Name & "'") Then Exit Sub

    ProcessAllTables transferExport
    WriteInformation "Finished Exports At: " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
End Sub

'@label:add-forms-codes
'@sub-title Copy module code into mapped UserForm components.
'@details
'Synchronises form modules by copying the code from each source module
'listed in forms tables into the corresponding UserForm component. A
'confirmation dialog is shown when prompts are enabled. After copying, the
'source module is removed from the VBProject to prevent duplication.
'Requires that host sheets are unlocked.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@throws ProjectError.ErrorUnexpectedState When a host sheet is protected.
'@Description("Copy module code into mapped forms according to the forms table.")
Public Sub AddFormsCodes()
    EnsureInitialised
    EnsureHostSheetsUnlocked
    SynchroniseFormsTables True
End Sub


'@section Deployment
'===============================================================================
'@description Full deployment workflow: protect, hide, and lock the workbook.

'@label:deploy
'@sub-title Protect worksheets, hide Dev assets, copy forms code, and set the deployment marker.
'@details
'Executes the full deployment workflow. First applies the default password
'from the Dev sheet, then synchronises form modules without confirmation.
'Every worksheet registered via AddHiddenSheet is set to xlSheetVeryHidden.
'Every worksheet registered via AddProtectedSheet is protected using the
'supplied IPasswords handler with the recorded permission flags. When
'lockWorkbook is True, the host workbook itself is locked. The code
'worksheet (if distinct from the Dev sheet) is hidden. Finally, a debug
'exit handler is installed and the workbook-level deployment flag is set
'to "Yes".
'@param pass IPasswords. The password handler used to apply protection to
'worksheets and optionally lock the workbook.
'@param lockWorkbook Optional Boolean. When True, the host workbook
'structure is locked after protecting sheets. Defaults to False.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been
'called or when pass is Nothing.
'@throws ProjectError.ErrorUnexpectedState When a host sheet is protected.
'@Description("Protect worksheets, hide Dev assets, copy forms code, and set deployment marker.")
Public Sub Deploy(ByVal pass As IPasswords, Optional ByVal lockWorkbook As Boolean = False)
    EnsureInitialised
    EnsureHostSheetsUnlocked

    If pass Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Passwords handler is required for deployment"
    End If

    Dim idx As Long
    Dim candidate As Variant
    Dim sheet As Worksheet

    ApplyDefaultPassword pass

    SynchroniseFormsTables False

    If Not this.hiddenSheets Is Nothing Then
        If this.hiddenSheets.Length > 0 Then
            For idx = this.hiddenSheets.LowerBound To this.hiddenSheets.UpperBound
                candidate = this.hiddenSheets.Item(idx)
                Set sheet = Nothing

                Select Case TypeName(candidate)
                    Case "Worksheet"
                        Set sheet = candidate
                    Case Else
                        Set sheet = TryResolveWorksheet(CStr(candidate))
                End Select

                If Not sheet Is Nothing Then
                    sheet.Visible = xlSheetVeryHidden
                End If
            Next idx
        End If
    End If

    If Not this.protectedSheets Is Nothing Then
        If this.protectedSheets.Length > 0 Then
            For idx = this.protectedSheets.LowerBound To this.protectedSheets.UpperBound
                Set candidate = this.protectedSheets.Item(idx)
                Set sheet = Nothing

                Select Case TypeName(candidate.Item(1))
                    Case "Worksheet"
                        Set sheet = candidate.Item(1)
                    Case Else
                        Set sheet = TryResolveWorksheet(CStr(candidate.Item(1)))
                End Select

                If Not sheet Is Nothing Then
                    pass.Protect sheet, candidate.Item(2), candidate.Item(3)
                End If
            Next idx
        End If
    End If

    If lockWorkbook Then
        pass.Protect this.hostwb
    End If

    If Not this.codeSheet Is Nothing Then
        If Not (this.codeSheet Is this.devSheet) Then
            this.codeSheet.Visible = xlSheetVeryHidden
        End If
    End If

    pass.EnsureDebugExitHandler this.hostwb
    AssignWorkbookName WORKBOOK_DEPLOYMENT_FLAG, EncodeNameValue("Yes")
End Sub

'@label:in-deployment
'@sub-title Check whether the workbook has already been deployed.
'@details
'Reads the workbook-level named range "inDeployment" and performs a
'case-insensitive comparison against "Yes". Returns True when the stored
'value matches, indicating that Deploy has already been executed. Returns
'False when the name does not exist or contains any other value.
'@return Boolean. True when the deployment flag equals "Yes".
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@Description("Return True when the workbook has already been deployed (flag equals 'Yes').")
Public Function InDeployment() As Boolean
    EnsureInitialised

    Dim definition As Name
    On Error Resume Next
        Set definition = this.hostwb.Names(WORKBOOK_DEPLOYMENT_FLAG)
    On Error GoTo 0

    If definition Is Nothing Then Exit Function

    'Workbook flag equals Yes when Deploy has run successfully.
    InDeployment = (StrComp(NameStringValue(definition), "Yes", vbTextCompare) = 0)
End Function


'@section Properties
'===============================================================================
'@description Read-only and read-write properties exposing the internal state.

'@label:dev-worksheet
'@prop-title Development worksheet reference.
'@details
'Returns the Dev worksheet that stores configuration named ranges such as
'folder paths and counter values. Set during Initialise and never changed
'afterwards.
'@return Worksheet. The development worksheet bound to this instance.
'@Description("Development worksheet storing configuration named ranges.")
Public Property Get DevWorksheet() As Worksheet
    Set DevWorksheet = this.devSheet
End Property

'@label:code-worksheet
'@prop-title Worksheet registered to host development tables.
'@details
'Returns the worksheet currently registered to host development ListObjects.
'May be Nothing when no code worksheet has been registered, in which case
'new tables are created on the Dev sheet itself.
'@return Worksheet. The code worksheet, or Nothing when none is registered.
'@Description("Worksheet currently registered to host development tables (may be Dev sheet).")
Public Property Get CodeWorksheet() As Worksheet
    Set CodeWorksheet = this.codeSheet
End Property

'@label:display-prompts
'@prop-title Whether confirmation prompts are displayed.
'@details
'Controls whether confirmation dialogs are shown before mutating operations
'such as importing, exporting, or adding tables. Defaults to True after
'Initialise. Set to False to suppress prompts during automated or batch
'workflows.
'@return Boolean. True when prompts are enabled.
'@Description("Whether user confirmation prompts are displayed before mutating operations.")
Public Property Get DisplayPrompts() As Boolean
    DisplayPrompts = this.displayPrompts
End Property

'@label:display-prompts-set
'@prop-title Set the confirmation prompt toggle.
'@param value Boolean. True to enable prompts, False to suppress them.
Public Property Let DisplayPrompts(ByVal value As Boolean)
    this.displayPrompts = value
End Property


'@section Worksheet Helpers
'===============================================================================
'@description Private helpers that manage worksheet discovery, validation,
'and collection initialisation.

'@label:ensure-collections
'@sub-title Initialise sheet tracking collections when not yet allocated.
'@Description("Initialise sheet tracking collections for hidden/protected lists.")
Private Sub EnsureCollections()
    If this.hiddenSheets Is Nothing Then
        Set this.hiddenSheets = New BetterArray
        this.hiddenSheets.LowerBound = 1
    End If

    If this.protectedSheets Is Nothing Then
        Set this.protectedSheets = New BetterArray
        this.protectedSheets.LowerBound = 1
    End If
End Sub

'@label:ensure-named-ranges
'@sub-title Validate that required Dev sheet named ranges exist.
'@details
'Checks that the module and class folder named ranges are present on the
'Dev sheet. The tests folder range is optional and is only validated when
'its name definition already exists at the workbook level.
'@throws ProjectError.ElementNotFound When a required named range is missing.
'@Description("Validate Dev sheet named ranges for module/class/test folders.")
Private Sub EnsureNamedRanges()
    EnsureRangeExists this.devSheet, MODULE_CODES_NAME, required:=True
    EnsureRangeExists this.devSheet, CLASS_CODES_NAME, required:=True

    If Not TryNameDefinition(TEST_CODES_NAME) Is Nothing Then
        EnsureRangeExists this.devSheet, TEST_CODES_NAME, required:=False
    End If
End Sub

'@label:restore-code-worksheet
'@sub-title Restore a previously registered code worksheet from persisted state.
'@details
'Reads the worksheet-level named range CODE_SHEET_NAMEDEF. When a valid
'sheet name is found and the worksheet exists, it is set as the active
'code worksheet. Silently exits when no persisted name exists.
'@Description("Restore previously registered code worksheet if stored.")
Private Sub RestoreCodeWorksheet()
    Dim definition As Name
    Set definition = SheetNameDefinition(CODE_SHEET_NAMEDEF)

    If definition Is Nothing Then Exit Sub

    Dim sheetName As String
    sheetName = NameStringValue(definition)
    If LenB(sheetName) = 0 Then Exit Sub

    Dim sheet As Worksheet
    On Error Resume Next
        Set sheet = this.hostwb.Worksheets(sheetName)
    On Error GoTo 0

    If sheet Is Nothing Then Exit Sub

    Set this.codeSheet = sheet
    AssignSheetName CODE_SHEET_NAMEDEF, EncodeNameValue(sheet.Name)
End Sub

'@label:ensure-worksheet
'@sub-title Resolve a worksheet by name, raising on empty name or missing sheet.
'@details
'Trims the supplied name and attempts to resolve it in the host workbook.
'Raises InvalidArgument when the name is empty and ElementNotFound when no
'matching worksheet exists. The purpose parameter is included in error
'messages to help callers identify the context.
'@param sheetName String. Name of the worksheet to resolve.
'@param purpose String. Label included in error messages for context.
'@return Worksheet. The resolved worksheet reference.
'@throws ProjectError.InvalidArgument When sheetName is empty.
'@throws ProjectError.ElementNotFound When the worksheet does not exist.
Private Function EnsureWorksheet(ByVal sheetName As String, ByVal purpose As String) As Worksheet
    Dim trimmedName As String
    trimmedName = Trim$(sheetName)

    If LenB(trimmedName) = 0 Then
        ThrowError ProjectError.InvalidArgument, _
                  "Worksheet name for " & purpose & " cannot be empty"
    End If

    Dim sheet As Worksheet
    On Error Resume Next
        Set sheet = this.hostwb.Worksheets(trimmedName)
    On Error GoTo 0

    If sheet Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Worksheet '" & trimmedName & "' not found in workbook '" & this.hostwb.Name & "'"
    End If

    Set EnsureWorksheet = sheet
End Function

'@label:try-resolve-worksheet
'@sub-title Attempt to resolve a worksheet by name without raising.
'@param sheetName String. Name of the worksheet to look up.
'@return Worksheet. The resolved worksheet, or Nothing when not found.
Private Function TryResolveWorksheet(ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet

    If this.hostwb Is Nothing Then Exit Function

    On Error Resume Next
        Set sheet = this.hostwb.Worksheets(sheetName)
    On Error GoTo 0

    Set TryResolveWorksheet = sheet
End Function

'@label:ensure-range-exists
'@sub-title Validate that a named range exists on a worksheet.
'@param sheet Worksheet. The worksheet to check.
'@param rangeName String. The named range to locate.
'@param required Optional Boolean. When True, raises an error if the range
'is missing. Defaults to True.
'@throws ProjectError.ElementNotFound When required is True and the range
'is not found.
'@Description("Ensure a named range exists on a worksheet, raising when mandatory.")
Private Sub EnsureRangeExists(ByVal sheet As Worksheet, _
                              ByVal rangeName As String, _
                              Optional ByVal required As Boolean = True)
    Dim target As Range

    On Error Resume Next
        Set target = sheet.Range(rangeName)
    On Error GoTo 0

    If target Is Nothing And required Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & rangeName & "' is missing on worksheet '" & sheet.Name & "'"
    End If
End Sub

'@label:register-code-worksheet
'@sub-title Persist a worksheet as the active code table host.
'@param sheetName String. Name of the worksheet to register.
'@Description("Persist the supplied worksheet as the active code worksheet.")
Private Sub RegisterCodeWorksheet(ByVal sheetName As String)
    Dim sheet As Worksheet
    Set sheet = EnsureWorksheet(sheetName, "code")

    Set this.codeSheet = sheet
    AssignSheetName CODE_SHEET_NAMEDEF, EncodeNameValue(sheet.Name)
End Sub

'@label:resolve-table-host
'@sub-title Return the worksheet where new tables should be created.
'@return Worksheet. The code worksheet when registered, otherwise the Dev sheet.
'@Description("Return the worksheet where new development tables should be created.")
Private Function ResolveTableHost() As Worksheet
    If this.codeSheet Is Nothing Then
        Set ResolveTableHost = this.devSheet
    Else
        Set ResolveTableHost = this.codeSheet
    End If
End Function

'@label:table-host-sheets
'@sub-title Collect worksheets that may contain development tables.
'@details
'Returns a Collection containing the Dev sheet and, when registered and
'distinct, the code worksheet. Used by import/export and form
'synchronisation routines to iterate all possible table hosts.
'@return Collection. One or two Worksheet references.
'@Description("Return worksheets that may contain development tables for processing.")
Private Function TableHostSheets() As Collection
    Dim hosts As Collection
    Set hosts = New Collection

    hosts.Add this.devSheet

    If Not this.codeSheet Is Nothing Then
        If Not (this.codeSheet Is this.devSheet) Then
            hosts.Add this.codeSheet
        End If
    End If

    Set TableHostSheets = hosts
End Function

'@label:ensure-host-sheets-unlocked
'@sub-title Verify that candidate worksheets are not protected.
'@details
'Iterates every worksheet returned by TableHostSheets and raises an error
'when any of them has ProtectContents set to True. This guard prevents
'import, export, and deployment operations from failing mid-way due to
'sheet protection.
'@throws ProjectError.ErrorUnexpectedState When a host sheet is protected.
'@Description("Confirm candidate worksheets are unprotected before modification.")
Private Sub EnsureHostSheetsUnlocked()
    Dim sheet As Worksheet

    For Each sheet In TableHostSheets()
        If sheet.ProtectContents Then
            ThrowError ProjectError.ErrorUnexpectedState, _
                      "Worksheet '" & sheet.Name & "' is protected; unlock it before running this operation."
        End If
    Next sheet
End Sub

'@label:ensure-initialised
'@sub-title Guard public API calls when Development has not been initialised.
'@throws ProjectError.ObjectNotInitialized When Initialise has not been called.
'@Description("Guard public API calls when Development has not been initialised.")
Private Sub EnsureInitialised()
    If this.devSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Initialise must be called before using Development"
    End If
End Sub

'@section Code Transfer Helpers
'===============================================================================
'@description Private helpers that drive the import and export of VBA
'components listed in development tables.

'@label:process-all-tables
'@sub-title Iterate every development table and dispatch import or export.
'@details
'Reads folder settings from the Dev sheet, then walks every ListObject on
'each host worksheet. Each table is passed to ProcessListObject, which
'interprets the tag above the table to decide scope (modules or classes),
'target folder, and whether interfaces should be included.
'@param mode DevelopmentTransferMode. Whether to import or export.
'@Description("Iterate every development table and process it for import/export.")
Private Sub ProcessAllTables(ByVal mode As DevelopmentTransferMode)
    Dim folders As TFolderSettings
    folders = ResolveFolderSettings()

    Dim lo As ListObject
    Dim sheet As Worksheet

    For Each sheet In TableHostSheets()
        'Process tables from Dev sheet first, then any registered code worksheet.
        For Each lo In sheet.ListObjects
            ProcessListObject lo, mode, folders
        Next lo
    Next sheet
End Sub

'@label:resolve-folder-settings
'@sub-title Read module, class, and test folder paths from the Dev sheet.
'@return TFolderSettings. The resolved folder paths.
'@Description("Read module/class/test folder settings from the Dev sheet.")
Private Function ResolveFolderSettings() As TFolderSettings
    ResolveFolderSettings.modulesPath = ReadSheetValue(MODULE_CODES_NAME, required:=True)
    ResolveFolderSettings.classesPath = ReadSheetValue(CLASS_CODES_NAME, required:=True)
    ResolveFolderSettings.testsPath = ReadSheetValue(TEST_CODES_NAME, required:=False)
End Function

'@label:read-sheet-value
'@sub-title Return the trimmed string value of a named range on the Dev sheet.
'@details
'Locates the named range by first checking worksheet-level names, then
'falling back to workbook-level names. When the range cannot be resolved
'and required is True, an error is raised. Returns an empty string when
'the range is missing and required is False.
'@param rangeName String. The named range to read.
'@param required Optional Boolean. When True, raises an error if the range
'is missing. Defaults to True.
'@return String. The trimmed cell value, or empty string when not found.
'@throws ProjectError.ElementNotFound When required is True and the range
'is missing.
'@Description("Return the string value of a named range, enforcing presence when required.")
Private Function ReadSheetValue(ByVal rangeName As String, Optional ByVal required As Boolean = True) As String
    Dim target As Range

    Dim definition As Name
    Set definition = TryNameDefinition(rangeName)

    If definition Is Nothing Then
        If required Then
            ThrowError ProjectError.ElementNotFound, _
                      "Named range '" & rangeName & "' is missing on worksheet '" & this.devSheet.Name & "'"
        End If
        Exit Function
    End If

    On Error Resume Next
        Set target = definition.RefersToRange
    On Error GoTo 0

    If target Is Nothing Then
        If required Then
            ThrowError ProjectError.ElementNotFound, _
                      "Named range '" & rangeName & "' is missing on worksheet '" & this.devSheet.Name & "'"
        End If
        Exit Function
    End If

    ReadSheetValue = Trim$(CStr(target.Value))
End Function

'@label:process-list-object
'@sub-title Process a single ListObject for import or export.
'@details
'Reads the tag cell above the table to determine the scope (modules,
'classes, tests, or forms) and the base folder. Reads the folder cell two
'rows above the table to determine the subfolder. Combines both paths and
'verifies the directory exists on disk. Delegates to TransferListEntries
'for actual component-by-component processing. Forms-tagged tables are
'skipped here (handled by SynchroniseFormsTables instead). Returns True
'when at least one component was transferred.
'@param lo ListObject. The development table to process.
'@param mode DevelopmentTransferMode. Whether to import or export.
'@param folders TFolderSettings. The resolved folder paths.
'@return Boolean. True when at least one entry was processed.
'@Description("Process a single ListObject, importing or exporting its entries.")
Private Function ProcessListObject(ByVal lo As ListObject, _
                                   ByVal mode As DevelopmentTransferMode, _
                                   ByRef folders As TFolderSettings) As Boolean

    Dim tagValue As String
    tagValue = LCase$(ListTag(lo))

    Dim baseFolder As String
    Dim scope As DevelopmentTransferScope
    Dim includeInterface As Boolean

    'Interpret the tag written above the table to decide scope and destination.
    Select Case tagValue
        Case TAG_MODULES
            baseFolder = folders.modulesPath
            scope = transferModules

        Case TAG_CLASSES
            baseFolder = folders.classesPath
            scope = transferClasses
            includeInterface = True

        Case TAG_TESTS
            If LenB(folders.testsPath) = 0 Then Exit Function
            baseFolder = AppendPath(folders.testsPath, "modules")
            scope = transferModules

        Case TAG_TEST_CLASSES
            If LenB(folders.testsPath) = 0 Then Exit Function
            baseFolder = AppendPath(folders.testsPath, "classes")
            scope = transferClasses
            includeInterface = True

        Case TAG_FORMS
            Exit Function

        Case Else
            Exit Function
    End Select

    Dim subFolder As String
    subFolder = ListFolder(lo)

    Dim targetFolder As String
    targetFolder = AppendPath(baseFolder, subFolder)

    If LenB(targetFolder) = 0 Then Exit Function
    If Not DirectoryExists(targetFolder) Then Exit Function

    Dim processed As Boolean
    processed = TransferListEntries(lo, targetFolder, scope, mode, includeInterface)

    If processed Then
        AppendInformationLog targetFolder, scope, mode
        ProcessListObject = True
    End If
End Function

'@label:list-tag
'@sub-title Retrieve the scope tag stored in the cell above the table.
'@param lo ListObject. The target table.
'@return String. The trimmed tag value, or empty string when absent.
'@Description("Retrieve the tag stored above the table (scope identifier).")
Private Function ListTag(ByVal lo As ListObject) As String
    Dim tagCell As Range

    On Error Resume Next
        Set tagCell = lo.Range.Cells(0, 1)
    On Error GoTo 0

    If tagCell Is Nothing Then Exit Function
    ListTag = Trim$(CStr(tagCell.Value))
End Function

'@label:compare-tag
'@sub-title Case-insensitive comparison between a table tag and an expected value.
'@param lo ListObject. The table whose tag is compared.
'@param expected String. The expected tag text.
'@return Boolean. True when the tags match.
'@Description("Case-insensitive comparison between a table tag and expected text.")
Private Function CompareTag(ByVal lo As ListObject, ByVal expected As String) As Boolean
    CompareTag = (StrComp(LCase$(ListTag(lo)), LCase$(expected), vbTextCompare) = 0)
End Function

'@label:list-folder
'@sub-title Retrieve the subfolder path stored two rows above the table.
'@param lo ListObject. The target table.
'@return String. The trimmed subfolder path, or empty string when absent.
'@Description("Retrieve the folder path written above the table (relative folder).")
Private Function ListFolder(ByVal lo As ListObject) As String
    Dim folderCell As Range

    On Error Resume Next
        Set folderCell = lo.Range.Cells(-1, 1)
    On Error GoTo 0

    If folderCell Is Nothing Then Exit Function
    ListFolder = Trim$(CStr(folderCell.Value))
End Function

'@label:append-path
'@sub-title Combine a base path and subpath using the platform separator.
'@param basePath String. The root folder path.
'@param subPath String. The relative subfolder to append.
'@return String. The combined path.
'@Description("Combine base and sub-path using the platform path separator.")
Private Function AppendPath(ByVal basePath As String, ByVal subPath As String) As String
    If LenB(basePath) = 0 Then
        AppendPath = Trim$(subPath)
        Exit Function
    End If

    If LenB(subPath) = 0 Then
        AppendPath = Trim$(basePath)
        Exit Function
    End If

    Dim sep As String
    sep = Application.PathSeparator

    If Right$(basePath, 1) = sep Then
        AppendPath = basePath & subPath
    Else
        AppendPath = basePath & sep & subPath
    End If
End Function

'@label:directory-exists
'@sub-title Return True when the supplied directory exists on disk.
'@param pathValue String. The directory path to check.
'@return Boolean. True when the directory exists.
'@Description("Return True when the supplied directory exists on disk.")
Private Function DirectoryExists(ByVal pathValue As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(pathValue)

    If LenB(trimmed) = 0 Then Exit Function

    Do While Right$(trimmed, 1) = Application.PathSeparator
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    Loop

    DirectoryExists = (Dir$(trimmed, vbDirectory) <> vbNullString)
End Function

'@label:transfer-list-entries
'@sub-title Iterate table rows and import or export each listed component.
'@details
'Reads the first column for component names and, when includeInterface is
'True, the second column for a Yes/No interface flag. For each non-empty
'entry the component is transferred via TransferComponent. When the flag
'is "yes", the corresponding interface file (prefixed with "I") is also
'transferred. Returns True when at least one component was processed.
'@param lo ListObject. The development table containing component entries.
'@param destinationFolder String. The disk folder for the transfer.
'@param scope DevelopmentTransferScope. Whether entries are modules or classes.
'@param mode DevelopmentTransferMode. Whether to import or export.
'@param includeInterface Boolean. When True, also transfer interface files.
'@return Boolean. True when at least one entry was transferred.
'@Description("Iterate table rows and import/export each listed component.")
Private Function TransferListEntries(ByVal lo As ListObject, _
                                     ByVal destinationFolder As String, _
                                     ByVal scope As DevelopmentTransferScope, _
                                     ByVal mode As DevelopmentTransferMode, _
                                     ByVal includeInterface As Boolean) As Boolean

    Dim moduleColumn As Range

    On Error Resume Next
        Set moduleColumn = lo.ListColumns(1).DataBodyRange
    On Error GoTo 0

    If moduleColumn Is Nothing Then Exit Function

    Dim interfaceColumn As Range
    If includeInterface And lo.ListColumns.Count >= 2 Then
        On Error Resume Next
            Set interfaceColumn = lo.ListColumns(2).DataBodyRange
        On Error GoTo 0
    End If

    Dim rowIndex As Long
    Dim codeName As String
    Dim flagValue As String

    For rowIndex = 1 To moduleColumn.Rows.Count
        codeName = Trim$(CStr(moduleColumn.Cells(rowIndex, 1).Value))

        If LenB(codeName) > 0 Then
            TransferComponent codeName, destinationFolder, scope, mode
            TransferListEntries = True

            If includeInterface Then
                'For classes the second column stores Yes/No for interface duplication.
                flagValue = ResolveInterfaceFlag(interfaceColumn, rowIndex)
                If flagValue = "yes" Then
                    TransferComponent "I" & codeName, destinationFolder, scope, mode
                End If
            End If
        End If
    Next rowIndex
End Function

'@label:resolve-interface-flag
'@sub-title Read the Yes/No interface flag for a given table row.
'@param interfaceColumn Range. The second column range, or Nothing.
'@param rowIndex Long. The 1-based row index within the column.
'@return String. Lowercase "yes" or "no", or empty string when unavailable.
'@Description("Return lowercase 'yes'/'no' flag for interface duplication, handling missing cells.")
Private Function ResolveInterfaceFlag(ByVal interfaceColumn As Range, ByVal rowIndex As Long) As String
    If interfaceColumn Is Nothing Then Exit Function
    If rowIndex > interfaceColumn.Rows.Count Then Exit Function

    ResolveInterfaceFlag = LCase$(Trim$(CStr(interfaceColumn.Cells(rowIndex, 1).Value)))
End Function

'@label:transfer-component
'@sub-title Build the file path and delegate to import or export.
'@param componentName String. The VBA component name.
'@param destinationFolder String. The disk folder for the transfer.
'@param scope DevelopmentTransferScope. Determines the file extension.
'@param mode DevelopmentTransferMode. Whether to import or export.
'@Description("Import or export a single component based on mode and scope.")
Private Sub TransferComponent(ByVal componentName As String, _
                              ByVal destinationFolder As String, _
                              ByVal scope As DevelopmentTransferScope, _
                              ByVal mode As DevelopmentTransferMode)

    Dim filePath As String
    Dim sep As String

    sep = Application.PathSeparator

    Select Case scope
        Case transferModules
            filePath = destinationFolder & sep & componentName & ".bas"
        Case transferClasses
            filePath = destinationFolder & sep & componentName & ".cls"
        Case Else
            filePath = destinationFolder & sep & componentName
    End Select

    Select Case mode
        Case transferImport
            ImportComponent componentName, filePath
        Case transferExport
            ExportComponent componentName, filePath
    End Select
End Sub

'@label:import-component
'@sub-title Import a VBA component from disk, replacing any existing one.
'@details
'Checks that the file exists on disk, then removes any existing component
'with the same name from the VBProject so that VB attributes are refreshed
'on import. Delegates to VBComponents.Import to perform the actual import.
'@param componentName String. The VBA component name to replace.
'@param filePath String. The full disk path of the .bas or .cls file.
'@Description("Import a code component from disk, replacing any existing component.")
Private Sub ImportComponent(ByVal componentName As String, ByVal filePath As String)
    If Dir$(filePath) = vbNullString Then Exit Sub

    Dim vbProj As Object
    Dim components As Object

    Set vbProj = this.hostwb.VBProject
    Set components = vbProj.VBComponents

    On Error Resume Next
        'Remove any existing component so attributes are refreshed on import.
        components.Remove components(componentName)
    On Error GoTo 0

    components.Import filePath
End Sub

'@label:export-component
'@sub-title Export a VBA component to disk, overwriting any previous file.
'@details
'Looks up the component in the VBProject. When found, deletes any existing
'file at the target path and exports the component. Silently exits when
'the component does not exist in the project.
'@param componentName String. The VBA component name to export.
'@param filePath String. The full disk path for the exported file.
'@Description("Export a code component to disk, overwriting any previous file.")
Private Sub ExportComponent(ByVal componentName As String, ByVal filePath As String)
    Dim vbProj As Object
    Dim component As Object

    Set vbProj = this.hostwb.VBProject

    On Error Resume Next
        Set component = vbProj.VBComponents(componentName)
    On Error GoTo 0

    If component Is Nothing Then Exit Sub

    On Error Resume Next
        If Dir$(filePath) <> vbNullString Then Kill filePath
    On Error GoTo 0

    component.Export filePath
End Sub

'@label:append-information-log
'@sub-title Append a timestamped entry describing the transfer action performed.
'@param targetPath String. The folder path that was processed.
'@param scope DevelopmentTransferScope. Whether modules or classes were transferred.
'@param mode DevelopmentTransferMode. Whether an import or export was performed.
'@Description("Append a timestamped entry describing the import/export action performed.")
Private Sub AppendInformationLog(ByVal targetPath As String, _
                                 ByVal scope As DevelopmentTransferScope, _
                                 ByVal mode As DevelopmentTransferMode)

    Dim infoCell As Range
    Set infoCell = EnsureInformationCell()
    If infoCell Is Nothing Then Exit Sub

    Dim logCell As Range
    Set logCell = infoCell.Offset(9, 0)

    Do While LenB(Trim$(CStr(logCell.Value))) > 0
        Set logCell = logCell.Offset(1, 0)
        If logCell.Row >= infoCell.Worksheet.Rows.Count Then Exit Sub
    Loop

    Dim actionLabel As String
    Select Case mode
        Case transferImport: actionLabel = "Imported "
        Case transferExport: actionLabel = "Exported "
        Case Else: actionLabel = "Processed "
    End Select

    Dim scopeLabel As String
    Select Case scope
        Case transferModules: scopeLabel = "Modules using path: "
        Case transferClasses: scopeLabel = "Classes using path: "
        Case Else: scopeLabel = "Items using path: "
    End Select

    logCell.Value = Format$(Now(), "yyyy-mm-dd hh:mm:ss") & " - " & actionLabel & scopeLabel & targetPath
End Sub

'@label:write-information
'@sub-title Write a status message to the Dev sheet Informations cell.
'@param messageText String. The message to display.
'@Description("Write a status message in the Dev sheet informations cell.")
Private Sub WriteInformation(ByVal messageText As String)
    Dim infoCell As Range
    Set infoCell = EnsureInformationCell()

    If infoCell Is Nothing Then Exit Sub
    infoCell.Value = messageText
End Sub

'@label:ensure-information-cell
'@sub-title Locate the Dev sheet status cell from the Informations named range.
'@return Range. The target cell, or Nothing when the range is not defined.
'@Description("Locate the Dev sheet status cell (Informations named range).")
Private Function EnsureInformationCell() As Range
    Dim infoCell As Range

    On Error Resume Next
        Set infoCell = this.devSheet.Range(INFO_RANGE)
    On Error GoTo 0

    Set EnsureInformationCell = infoCell
End Function

'@section Form Synchronisation
'===============================================================================
'@description Private helpers that copy code from standard modules into
'UserForm components based on forms tables.

'@label:synchronise-forms-tables
'@sub-title Enumerate forms tables and copy code into mapped UserForm components.
'@details
'Walks every ListObject on each host worksheet and identifies those tagged
'as "general form modules". For each matching table, ProcessFormsTable is
'called to copy module code into the corresponding form component. When
'requireConfirmation is True, a dialog is shown before proceeding.
'@param requireConfirmation Boolean. When True, prompts the user first.
'@Description("Enumerate all forms tables and optionally prompt before synchronising code.")
Private Sub SynchroniseFormsTables(ByVal requireConfirmation As Boolean)
    If requireConfirmation Then
        If Not RequestConfirmation("synchronise form modules with mapped code modules") Then Exit Sub
    End If

    Dim sheet As Worksheet
    Dim lo As ListObject

    For Each sheet In TableHostSheets()
        For Each lo In sheet.ListObjects
            If CompareTag(lo, TAG_FORMS) Then
                ProcessFormsTable lo
            End If
        Next lo
    Next sheet
End Sub

'@label:apply-default-password
'@sub-title Read and apply the default password from the Dev sheet.
'@details
'Checks for a named range "DefaultPassword" on the Dev sheet. When found
'and non-empty, passes its value to the IPasswords handler via
'NewDebugingPassword. Silently exits when no default password is configured.
'@param passwords IPasswords. The password handler to configure.
Private Sub ApplyDefaultPassword(ByVal passwords As IPasswords)
    Dim defaultCell As Range
    Dim defaultValue As String

    If passwords Is Nothing Then Exit Sub
    If this.devSheet Is Nothing Then Exit Sub

    On Error Resume Next
        Set defaultCell = this.devSheet.Range("DefaultPassword")
    On Error GoTo 0

    If defaultCell Is Nothing Then Exit Sub

    defaultValue = Trim$(CStr(defaultCell.Value))
    If LenB(defaultValue) = 0 Then Exit Sub

    On Error Resume Next
        passwords.NewDebugingPassword defaultValue
    On Error GoTo 0
End Sub

'@label:process-forms-table
'@sub-title Copy code for each module/form pair defined in a forms ListObject.
'@details
'Reads the first column for source module names and the second column for
'target form names. For each non-empty pair, CopyModuleCode transfers the
'code from the module into the form component and removes the source.
'@param lo ListObject. The forms mapping table to process.
'@Description("Copy code for each module/form pair defined in the supplied ListObject.")
Private Sub ProcessFormsTable(ByVal lo As ListObject)
    Dim moduleColumn As Range
    Dim formColumn As Range

    On Error Resume Next
        Set moduleColumn = lo.ListColumns(1).DataBodyRange
        Set formColumn = lo.ListColumns(2).DataBodyRange
    On Error GoTo 0

    If moduleColumn Is Nothing Then Exit Sub
    If formColumn Is Nothing Then Exit Sub

    Dim idx As Long
    Dim moduleName As String
    Dim targetName As String

    For idx = 1 To moduleColumn.Rows.Count
        moduleName = Trim$(CStr(moduleColumn.Cells(idx, 1).Value))
        targetName = Trim$(CStr(formColumn.Cells(idx, 1).Value))

        If LenB(moduleName) > 0 And LenB(targetName) > 0 Then
            CopyModuleCode moduleName, targetName
        End If
    Next idx
End Sub

'@label:copy-module-code
'@sub-title Copy the VBIDE code from one component into another.
'@details
'Reads all lines from the source component CodeModule, clears the target
'component CodeModule, and writes the source content into it. The source
'component is then removed from the VBProject. Raises ElementNotFound when
'either the source or target component does not exist.
'@param sourceName String. Name of the VBA component to copy from.
'@param targetName String. Name of the VBA component to copy into.
'@throws ProjectError.ElementNotFound When the source or target component
'is not found in the VBProject.
'@Description("Copy the VBIDE code from one component into another.")
Private Sub CopyModuleCode(ByVal sourceName As String, ByVal targetName As String)
    Dim vbProj As Object
    Dim sourceComponent As Object
    Dim targetComponent As Object

    Set vbProj = this.hostwb.VBProject

    On Error Resume Next
        Set sourceComponent = vbProj.VBComponents(sourceName)
        Set targetComponent = vbProj.VBComponents(targetName)
    On Error GoTo 0

    If sourceComponent Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Source module '" & sourceName & "' not found in workbook '" & this.hostwb.Name & "'"
    End If

    If targetComponent Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Target module '" & targetName & "' not found in workbook '" & this.hostwb.Name & "'"
    End If

    Dim codeContent As String
    With sourceComponent.CodeModule
        If .CountOfLines > 0 Then
            codeContent = .Lines(1, .CountOfLines)
        Else
            codeContent = vbNullString
        End If
    End With

    With targetComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If LenB(codeContent) > 0 Then
            .AddFromString codeContent
        End If
    End With

    On Error Resume Next
        vbProj.VBComponents.Remove sourceComponent
    On Error GoTo 0

End Sub


'@section ListObject Builders
'===============================================================================
'@description Private helpers that create, format, and name development
'tables on the code worksheet.

'@label:create-list-object
'@sub-title Create a table with standardised headers, tag, and naming.
'@details
'Resolves the host worksheet, calculates the next available horizontal
'anchor, writes the folder and tag cells above the table, creates a
'ListObject from a two-row range (header + one data row), applies project
'formatting, generates a unique name via the persisted counter, and
'optionally adds Yes/No data validation to the second column.
'@param headers Variant. Array of column header strings.
'@param tagValue String. The scope tag written above the table.
'@param counterName String. The worksheet-level name storing the counter.
'@param prefix String. The prefix for the generated ListObject name.
'@param addYesNoValidation Optional Boolean. When True, applies a Yes/No
'dropdown to the second column. Defaults to False.
'@return ListObject. The newly created and formatted table.
'@Description("Create a table with the supplied headers, tag, and standardised naming.")
Private Function CreateListObject(ByVal headers As Variant, _
                                  ByVal tagValue As String, _
                                  ByVal counterName As String, _
                                  ByVal prefix As String, _
                                  Optional ByVal addYesNoValidation As Boolean = False) As ListObject

    Dim sheet As Worksheet
    Set sheet = ResolveTableHost()

    Dim columnCount As Long
    columnCount = UBound(headers) - LBound(headers) + 1

    Dim anchor As Range
    'Pick the next free block to the right of existing tables so layouts grow horizontally.
    Set anchor = NextAnchorRange(sheet, columnCount)

    Dim folderCell As Range
    Dim tagCell As Range
    Dim headerStart As Range
    Dim headerRange As Range
    Dim Lo As ListObject
    Dim listName As String

    Set folderCell = anchor.Cells(1, 1)
    Set tagCell = folderCell.Offset(1, 0)
    Set headerStart = folderCell.Offset(2, 0)
    Set headerRange = headerStart.Resize(2, columnCount)

    headerRange.Clear
    folderCell.Clear
    tagCell.Value = tagValue

    Set Lo = sheet.ListObjects.Add(xlSrcRange, headerRange, , xlYes)

    WriteHeaders Lo, headers
    ApplyLoFormat Lo
    listName = NextListObjectName(sheet, prefix, counterName)
    Lo.Name = listName

    If addYesNoValidation Then
        ApplyYesNoValidation Lo
    End If

    Set CreateListObject = Lo
End Function


'@label:apply-lo-format
'@sub-title Apply project formatting to a development ListObject.
'@param Lo ListObject. The table to format.
Private Sub ApplyLoFormat(ByVal Lo As Listobject)
    Lo.DataBodyRange.ClearContents
    Lo.TableStyle = "TableStyleLight11"
    Lo.HeaderRowRange.Interior.Color = RGB(11, 48, 64)
    Lo.HeaderRowRange.Cells(0, 1).Font.Color = RGB(56, 37, 35)
    Lo.HeaderRowRange.Cells(-1, 1).Font.Color = RGB(56, 37, 35)
     Lo.HeaderRowRange.Cells(0, 1).HorizontalAlignment = xlHAlignCenter
    Lo.HeaderRowRange.Cells(-1, 1).HorizontalAlignment = xlHAlignCenter
    Lo.HeaderRowRange.Font.Color = vbWhite
    Lo.HeaderRowRange.Font.Bold = True
    Lo.HeaderRowRange.EntireColumn.AutoFit
    Lo.HeaderRowRange.EntireRow.RowHeight = 20
End Sub

'@label:write-headers
'@sub-title Populate a table header row with the provided field names.
'@param lo ListObject. The table whose headers are written.
'@param headers Variant. Array of header strings.
'@Description("Populate table header row with provided field names.")
Private Sub WriteHeaders(ByVal lo As ListObject, ByVal headers As Variant)
    Dim idx As Long
    Dim headerRow As Range

    Set headerRow = lo.HeaderRowRange

    For idx = LBound(headers) To UBound(headers)
        headerRow.Cells(1, idx - LBound(headers) + 1).Value = headers(idx)
    Next idx
End Sub

'@label:apply-yes-no-validation
'@sub-title Apply a Yes/No dropdown validation to the second column.
'@param lo ListObject. The table to configure.
'@Description("Apply Yes/No validation to the second column when applicable.")
Private Sub ApplyYesNoValidation(ByVal lo As ListObject)
    If lo.ListColumns.Count < 2 Then Exit Sub

    Dim target As Range
    Set target = lo.ListColumns(2).DataBodyRange

    If target Is Nothing Then
        Set target = lo.ListColumns(2).Range.Offset(1, 0).Resize(1, 1)
    End If

    target.Validation.Delete
    target.Validation.Add Type:=xlValidateList, _
                          AlertStyle:=xlValidAlertStop, _
                          Operator:=xlBetween, _
                          Formula1:=YES_NO_LIST
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub

'@label:next-anchor-range
'@sub-title Find the next available anchor position for a table.
'@details
'Scans existing ListObjects and occupied cells on the header row to find
'the rightmost occupied column. The anchor is placed two columns to the
'right of this boundary. When the sheet lacks sufficient columns for the
'requested table width, an error is raised.
'@param sheet Worksheet. The worksheet to scan.
'@param columnCount Long. Number of columns needed for the new table.
'@return Range. The top-left anchor cell for the new table.
'@throws ProjectError.ErrorUnexpectedState When there is not enough space.
'@Description("Find the next available anchor position for a table, growing to the right.")
Private Function NextAnchorRange(ByVal sheet As Worksheet, ByVal columnCount As Long) As Range
    Dim baseCell As Range
    Set baseCell = sheet.Range(DEFAULT_ANCHOR)

    Dim headerRow As Long
    headerRow = baseCell.Row + 2
    'Existing tables advertise their header row two rows below anchor (tag and folder rows occupy first two).

    Dim rightMost As Long
    rightMost = baseCell.Column - 2

    Dim lo As ListObject
    For Each lo In sheet.ListObjects
        If lo.HeaderRowRange.Row = headerRow Then
            Dim candidate As Long
            candidate = lo.Range.Column + lo.ListColumns.Count - 1
            If candidate > rightMost Then
                rightMost = candidate
            End If
        End If
    Next lo

    Dim endCell As Range
    On Error Resume Next
        Set endCell = sheet.Cells(headerRow, sheet.Columns.Count).End(xlToLeft)
    On Error GoTo 0

    If Not endCell Is Nothing Then
        If endCell.Column >= baseCell.Column Then
            If LenB(CStr(endCell.Value)) > 0 Then
                If endCell.Column > rightMost Then
                    rightMost = endCell.Column
                End If
            ElseIf endCell.Column > baseCell.Column Then
                If endCell.Column > rightMost Then
                    rightMost = endCell.Column
                End If
            End If
        End If
    End If

    Dim targetColumn As Long
    If rightMost < baseCell.Column Then
        targetColumn = baseCell.Column
    Else
        targetColumn = rightMost + 2
    End If

    If targetColumn + columnCount - 1 > sheet.Columns.Count Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Insufficient space to add another table on worksheet '" & sheet.Name & "'"
    End If

    Set NextAnchorRange = sheet.Cells(baseCell.Row, targetColumn)
End Function

'@label:next-list-object-name
'@sub-title Generate a unique ListObject name using the persisted counter.
'@param sheet Worksheet. The worksheet hosting the table.
'@param prefix String. The name prefix (e.g. "ClassesLo").
'@param counterName String. The worksheet-level name storing the counter.
'@return String. A unique name such as "ClassesLo3".
'@Description("Generate a unique ListObject name using the persisted counter.")
Private Function NextListObjectName(ByVal sheet As Worksheet, _
                                    ByVal prefix As String, _
                                    ByVal counterName As String) As String
    Dim nextValue As Long
    Dim candidate As String

    nextValue = ReadCounter(counterName)

    Do
        nextValue = nextValue + 1
        candidate = prefix & CStr(nextValue)
    Loop While HasListObject(sheet, candidate)

    UpdateCounter counterName, nextValue
    NextListObjectName = candidate
End Function

'@label:has-list-object
'@sub-title Check whether a worksheet already defines a ListObject with the given name.
'@param sheet Worksheet. The worksheet to search.
'@param listName String. The ListObject name to look for.
'@return Boolean. True when a matching ListObject exists.
'@Description("Check whether the worksheet already defines a ListObject with the given name.")
Private Function HasListObject(ByVal sheet As Worksheet, ByVal listName As String) As Boolean
    Dim lo As ListObject

    For Each lo In sheet.ListObjects
        If StrComp(lo.Name, listName, vbTextCompare) = 0 Then
            HasListObject = True
            Exit Function
        End If
    Next lo
End Function


'@section Counter Helpers
'===============================================================================
'@description Private helpers that read and write persisted numeric
'counters stored as worksheet-level names.

'@label:read-counter
'@sub-title Read the persisted numeric counter stored as a worksheet name.
'@param nameId String. The name definition identifier.
'@return Long. The current counter value, or 0 when not found.
'@Description("Read the persisted numeric counter stored as a worksheet name.")
Private Function ReadCounter(ByVal nameId As String) As Long
    Dim definition As Name
    Set definition = SheetNameDefinition(nameId)

    If definition Is Nothing Then Exit Function
    ReadCounter = NameNumericValue(definition)
End Function

'@label:update-counter
'@sub-title Persist the next counter value back to the Dev sheet.
'@param nameId String. The name definition identifier.
'@param value Long. The counter value to store.
'@Description("Persist the next counter value back to the Dev sheet.")
Private Sub UpdateCounter(ByVal nameId As String, ByVal value As Long)
    AssignSheetName nameId, "=" & CStr(value)
End Sub

'@label:sheet-name-definition
'@sub-title Locate a worksheet-level name definition matching the supplied id.
'@param nameId String. The name to look up (ignoring sheet prefix).
'@return Name. The matching definition, or Nothing when not found.
'@Description("Locate a worksheet-level name definition matching the supplied id.")
Private Function SheetNameDefinition(ByVal nameId As String) As Name
    Dim idx As Long

    For idx = 1 To this.devSheet.Names.Count
        If CompareName(this.devSheet.Names(idx), nameId) Then
            Set SheetNameDefinition = this.devSheet.Names(idx)
            Exit Function
        End If
    Next idx
End Function

'@label:try-name-definition
'@sub-title Locate a name definition at worksheet or workbook level.
'@details
'Checks worksheet-level names first via SheetNameDefinition, then falls
'back to the workbook-level Names collection.
'@param nameId String. The name to look up.
'@return Name. The matching definition, or Nothing when not found.
Private Function TryNameDefinition(ByVal nameId As String) As Name
    Dim definition As Name

    Set definition = SheetNameDefinition(nameId)
    If Not definition Is Nothing Then
        Set TryNameDefinition = definition
        Exit Function
    End If

    On Error Resume Next
        Set definition = this.hostwb.Names(nameId)
    On Error GoTo 0

    Set TryNameDefinition = definition
End Function

'@label:assign-sheet-name
'@sub-title Assign or replace a worksheet-level name with the supplied formula.
'@param nameId String. The name definition identifier.
'@param refersTo String. The formula or value to store.
'@Description("Assign or replace a worksheet-level name with the supplied formula.")
Private Sub AssignSheetName(ByVal nameId As String, ByVal refersTo As String)
    RemoveSheetName nameId
    this.devSheet.Names.Add Name:=nameId, RefersTo:=refersTo
End Sub

'@label:remove-sheet-name
'@sub-title Remove a worksheet-level name if it already exists.
'@param nameId String. The name definition to delete.
'@Description("Remove a worksheet-level name if it already exists.")
Private Sub RemoveSheetName(ByVal nameId As String)
    Dim idx As Long

    For idx = this.devSheet.Names.Count To 1 Step -1
        If CompareName(this.devSheet.Names(idx), nameId) Then
            this.devSheet.Names(idx).Delete
        End If
    Next idx
End Sub

'@label:assign-workbook-name
'@sub-title Create or replace a workbook-level name with the supplied formula.
'@param nameId String. The name definition identifier.
'@param refersTo String. The formula or value to store.
'@Description("Create or replace a workbook-level name with the supplied formula.")
Private Sub AssignWorkbookName(ByVal nameId As String, ByVal refersTo As String)
    RemoveWorkbookName nameId
    this.hostwb.Names.Add Name:=nameId, RefersTo:=refersTo
End Sub

'@label:remove-workbook-name
'@sub-title Delete a workbook-level name when present.
'@param nameId String. The name definition to delete.
'@Description("Delete a workbook-level name when present.")
Private Sub RemoveWorkbookName(ByVal nameId As String)
    Dim idx As Long

    For idx = this.hostwb.Names.Count To 1 Step -1
        If CompareWorkbookName(this.hostwb.Names(idx), nameId) Then
            this.hostwb.Names(idx).Delete
        End If
    Next idx
End Sub

'@label:compare-workbook-name
'@sub-title Case-insensitive comparison for workbook-level names.
'@param definition Name. The name definition to compare.
'@param targetName String. The expected name string.
'@return Boolean. True when names match.
'@Description("Case-insensitive comparison for workbook-level names.")
Private Function CompareWorkbookName(ByVal definition As Name, ByVal targetName As String) As Boolean
    CompareWorkbookName = (StrComp(definition.Name, targetName, vbTextCompare) = 0)
End Function

'@label:compare-name
'@sub-title Case-insensitive comparison for worksheet-level names, ignoring sheet prefix.
'@details
'Worksheet-level names are stored as "SheetName!DefinitionName". This
'method strips the sheet prefix before comparing.
'@param definition Name. The name definition to compare.
'@param targetName String. The expected name string (without prefix).
'@return Boolean. True when names match.
'@Description("Case-insensitive comparison for worksheet-level names, ignoring sheet prefix.")
Private Function CompareName(ByVal definition As Name, ByVal targetName As String) As Boolean
    Dim suffix As String
    Dim delimiter As Long

    delimiter = InStrRev(definition.Name, "!")
    If delimiter > 0 Then
        suffix = Mid$(definition.Name, delimiter + 1)
    Else
        suffix = definition.Name
    End If

    CompareName = (StrComp(suffix, targetName, vbTextCompare) = 0)
End Function

'@section Name Utilities
'===============================================================================
'@description Private helpers for encoding, decoding, and reading values
'from Excel Name definitions.

'@label:name-numeric-value
'@sub-title Convert a name definition to a numeric value.
'@details
'Attempts to parse the RefersTo formula as a number. When the formula is
'not numeric, falls back to reading the value of the referred range.
'Returns 0 when neither approach yields a number.
'@param definition Name. The name definition to read.
'@return Long. The numeric value, or 0 when not parseable.
'@Description("Convert a name definition to a numeric value, using RefersTo or range value.")
Private Function NameNumericValue(ByVal definition As Name) As Long
    Dim refersTo As String
    Dim numericValue As Long

    refersTo = definition.RefersTo
    If LenB(refersTo) > 0 And Left$(refersTo, 1) = "=" Then
        refersTo = Mid$(refersTo, 2)
    End If

    If IsNumeric(refersTo) Then
        numericValue = CLng(refersTo)
    Else
        On Error Resume Next
            numericValue = CLng(definition.RefersToRange.Value)
        On Error GoTo 0
    End If

    NameNumericValue = numericValue
End Function

'@label:name-string-value
'@sub-title Return the decoded string value stored in a name definition.
'@param definition Name. The name definition to read.
'@return String. The decoded plain text value.
'@Description("Return the string value stored in a name definition.")
Private Function NameStringValue(ByVal definition As Name) As String
    NameStringValue = DecodeNameValue(definition.Value)
End Function

'@label:encode-name-value
'@sub-title Encode plain text as a quoted formula suitable for Name.RefersTo.
'@param value String. The text to encode.
'@return String. The encoded formula string (e.g. ="Hello").
'@Description("Encode plain text as a quoted formula suitable for Name.RefersTo.")
Private Function EncodeNameValue(ByVal value As String) As String
    EncodeNameValue = "=""" & Replace(value, """", """""") & """"
End Function

'@label:decode-name-value
'@sub-title Decode a Name.RefersTo formula to raw string content.
'@details
'Strips the leading "=" and surrounding double quotes, then unescapes
'embedded doubled quotes back to single quotes.
'@param refersTo String. The raw Name.Value or Name.RefersTo string.
'@return String. The decoded plain text.
'@Description("Decode a Name.RefersTo formula to raw string content.")
Private Function DecodeNameValue(ByVal refersTo As String) As String
    Dim text As String
    text = refersTo

    If Len(text) > 0 And Left$(text, 1) = "=" Then
        text = Mid$(text, 2)
    End If

    If Len(text) >= 2 Then
        If Left$(text, 1) = """" And Right$(text, 1) = """" Then
            text = Mid$(text, 2, Len(text) - 2)
        End If
    End If

    DecodeNameValue = Replace(text, """""", """")
End Function


'@section Registration Helpers
'===============================================================================
'@description Private helper for registering worksheet entries into the
'protected sheets collection.

'@label:register-sheet
'@sub-title Register a worksheet name with permission flags in a BetterArray.
'@details
'Wraps the sheet name, allowShape, and allowDeletingRows values into a
'BetterArray triplet and pushes it onto the registry. Used by
'AddProtectedSheet to build the protectedSheets collection that Deploy
'iterates during deployment.
'@param registry BetterArray. The collection to append to.
'@param sheetName String. The worksheet name to register.
'@param allowShape Optional Boolean. Whether shapes remain editable. Defaults to True.
'@param allowDeletingRows Optional Boolean. Whether row deletion is permitted. Defaults to True.
'@Description("Register a worksheet name in the provided BetterArray if not already tracked.")
Private Sub RegisterSheet(ByVal registry As BetterArray, ByVal sheetName As String, _
                          Optional ByVal allowShape As Boolean = True, _
                          Optional ByVal allowDeletingRows As Boolean = True)

    If registry Is Nothing Then Exit Sub

    Dim sheetInfo As BetterArray
    Dim trimmed As String

    trimmed = Trim$(sheetName)
    If LenB(trimmed) = 0 Then Exit Sub

    Set sheetInfo = New BetterArray
    sheetInfo.LowerBound = 1
    sheetInfo.Push sheetName, allowShape, allowDeletingRows

    'Adding the sheet informations into the registry
    registry.Push sheetInfo

End Sub


'@section Prompts
'===============================================================================
'@description User confirmation dialog management.

'@label:request-confirmation
'@sub-title Display a confirmation dialog when prompts are enabled.
'@details
'When displayPrompts is False, returns True immediately without showing a
'dialog. Otherwise displays a Yes/No message box. Returns True only when
'the user selects Yes.
'@param action String. A verb phrase describing the action being confirmed.
'@return Boolean. True when the user confirms or prompts are disabled.
'@Description("Display a confirmation dialog when prompts are enabled.")
Private Function RequestConfirmation(ByVal action As String) As Boolean
    If Not this.displayPrompts Then
        RequestConfirmation = True
        Exit Function
    End If

    Dim prompt As String
    prompt = "Are you sure you want to " & action & "?"

    RequestConfirmation = (MsgBox(prompt, vbYesNo + vbQuestion, CONFIRM_TITLE) = vbYes)
End Function


'@section Error Handling
'===============================================================================
'@description Centralised error-raising helper.

'@label:throw-error
'@sub-title Raise a ProjectError-based exception with a prefixed message.
'@param errNumber ProjectError. The error code to raise.
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
'@Description("Raise a project-scoped runtime error with prefixed message.")
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, "Error: " & message
End Sub


'@section Interface Implementation
'===============================================================================
'@description Delegated members that satisfy the IDevelopment contract.
'See the corresponding Public members above for full documentation.

Private Sub IDevelopment_Initialise(ByVal devSheet As Worksheet, Optional ByVal codeSheet As Worksheet)
    Initialise devSheet, codeSheet
End Sub

Private Property Get IDevelopment_DevWorksheet() As Worksheet
    Set IDevelopment_DevWorksheet = DevWorksheet
End Property

Private Property Get IDevelopment_CodeWorksheet() As Worksheet
    Set IDevelopment_CodeWorksheet = CodeWorksheet
End Property

Private Property Get IDevelopment_DisplayPrompts() As Boolean
    IDevelopment_DisplayPrompts = DisplayPrompts
End Property

Private Property Let IDevelopment_DisplayPrompts(ByVal value As Boolean)
    DisplayPrompts = value
End Property

Private Function IDevelopment_AddCodeSheets(ByVal sheetName As String) As Worksheet
    Set IDevelopment_AddCodeSheets = AddCodeSheets(sheetName)
End Function

Private Function IDevelopment_AddClassTable(Optional ByVal includeTests As Boolean = False) As ListObject
    Set IDevelopment_AddClassTable = AddClassTable(includeTests)
End Function

Private Function IDevelopment_AddModuleTable(Optional ByVal includeTests As Boolean = False) As ListObject
    Set IDevelopment_AddModuleTable = AddModuleTable(includeTests)
End Function

Private Function IDevelopment_AddFormsTable() As ListObject
    Set IDevelopment_AddFormsTable = AddFormsTable()
End Function

Private Function IDevelopment_AddTestTable() As ListObject
    Set IDevelopment_AddTestTable = AddTestTable()
End Function

Private Sub IDevelopment_AddHiddenSheet(ByVal sheetName As String)
    AddHiddenSheet sheetName
End Sub

Private Sub IDevelopment_AddProtectedSheet(ByVal sheetName As String, _
                                           Optional ByVal allowShapes As Boolean = True, _
                                           Optional ByVal allowDeletingRows As Boolean = True)

    AddProtectedSheet sheetName, allowShapes, allowDeletingRows

End Sub

Private Sub IDevelopment_ImportAll()
    ImportAll
End Sub

Private Sub IDevelopment_ExportAll()
    ExportAll
End Sub

Private Sub IDevelopment_AddFormsCodes()
    AddFormsCodes
End Sub

Private Sub IDevelopment_Deploy(ByVal passwords As IPasswords, Optional ByVal lockWorkbook As Boolean = False)
    Deploy passwords, lockWorkbook
End Sub

Private Function IDevelopment_InDeployment() As Boolean
    IDevelopment_InDeployment = InDeployment
End Function
