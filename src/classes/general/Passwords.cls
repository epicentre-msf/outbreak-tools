VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Passwords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Lock/unlock workbooks and worksheets"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Lock/unlock workbooks and worksheets")
'@depends Checking, ProjectError, TranslationObject
'@IgnoreModule ProcedureNotUsed, ReadOnlyPropertyAssignment, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@class Passwords
'@description
'Manages workbook/worksheet protection, key generation, debug-mode
'toggling, and import/export of password metadata. Uses TypeName
'dispatch to handle Workbook, Worksheet, and String target arguments
'through a single Protect/UnProtect entry point.
'@details
'This class is on the hot path of a whole build, so it reads the worksheet as
'little as it can. The debugging password and the debug-mode flag are each read
'once and kept on the instance; ResetCaches drops them, and ResetCaches runs
'when the sheet is bound and after ImportFrom. The protection table is read as
'one block and matched in memory.
'
'Excel state belongs to the caller. This class never touches
'Application.ScreenUpdating, Calculation or EnableEvents.
'@depends Checking, ProjectError, TranslationObject
'@version 1.1 (2026-07-28)


'@section State and Constants
'===============================================================================
'Tag identifiers for the named ranges this class reads. The enum is Private:
'nothing outside the class names a tag. Parameters take Byte, which is the
'macOS-safe way to carry an enum value.
Private Enum PasswordTag
  PasswordTagDebug = 1
  PasswordTagPublic = 2
  PasswordTagPrivate = 3
  PasswordTagVersion = 4
  PasswordTagMode = 5
  PasswordTagLabPublic = 6
  PasswordTagLabPrivate = 7
End Enum

Private Type TPasswords
    sh As Worksheet                          'Password worksheet backing the object
    wb As Workbook                           'Cached workbook reference
    keyTable As ListObject                   'Cached list object holding keys
    protectedTable As ListObject             'Cached list object holding protection metadata
    debugPassword As String                  'Cached debugging password
    debugPasswordLoaded As Boolean           'Whether the debugging password was read
    debugMode As Boolean                     'Cached debug-mode flag
    debugModeLoaded As Boolean               'Whether the debug-mode flag was read
    hasChecks As Boolean                     'Whether checkings were emitted
    checks As Checking                       'Checking aggregation object
    checkCounter As Long                     'Incremental counter for checking keys
    nm As String                             'Name of the password worksheet
    displayPrompts As Boolean                'Whether UI prompts should be displayed
    lastPrivatePrompt As String              'Most recent private key prompt text
    lastPrivateTitle As String               'Most recent private key prompt title
    sealed As Boolean                        'Whether creation-time setters are sealed
End Type

Private this As TPasswords

Private Const CLASSNAME As String = "Passwords"
Private Const TABLEKEYS As String = "T_keys"
Private Const TABLEPROTECTED As String = "T_ProtectedSheets"
Private Const NAMEDEBUGPASSWORD As String = "RNG_DebuggingPassword"
Private Const NAMEPUBLICKEY As String = "RNG_PublicKey"
Private Const NAMELABPUBLICKEY As String = "RNG_LabPublicKey"
Private Const NAMEPRIVATEKEY As String = "RNG_PrivateKey"
Private Const NAMELABPRIVATEKEY As String = "RNG_LabPrivateKey"
Private Const NAMEDEBUGMODE As String = "RNG_DebugMode"
Private Const NAMEVERSION As String = "RNG_Version"
Private Const NAMEPROTECTEDSHEETS As String = "Passwords_ProtectedSheets"
Private Const DEFAULTBOOLYES As String = "yes"
Private Const DEFAULTBOOLNO As String = "no"
Private Const CHECKINGCONTEXT As String = "Passwords object trace"
Private Const PROTECTIONCOLUMNS As Long = 3



'@section Core Accessors
'===============================================================================
'Factory, accessors, and stored-value retrieval.

'@label:create
'@sub-title Create an instance of password
'@details
'Factory returning a password handler bound to the provided worksheet. The
'sheet is validated first: both tables and the five core named ranges must be
'there. The instance is sealed, so the setup setter can no longer be written.
'@param sh Worksheet. Worksheet hosting password tables and named ranges.
'@return Passwords. Fully initialised password handler.
'@export
Public Function Create(ByVal sh As Worksheet) As Passwords

    ValidateSheet sh

    Dim newPasswords As Passwords
    Set newPasswords = New Passwords

    Set newPasswords.PasswordSheet = sh
    newPasswords.Seal
    Set Create = newPasswords

End Function

'@label:password-sheet-set
'@prop-title Assign the password worksheet and reset caches
'@param sh Worksheet. The password worksheet to bind.
Public Property Set PasswordSheet(ByVal sh As Worksheet)
    GuardNotSealed "PasswordSheet"
    Set this.sh = sh
    Set this.wb = sh.Parent
    this.nm = sh.Name()
    this.displayPrompts = True
    this.lastPrivatePrompt = vbNullString
    this.lastPrivateTitle = vbNullString
    ResetCaches
End Property

'@label:name
'@prop-title Name of the password worksheet
'@details
'The name is taken once, when the sheet is bound. Renaming the sheet afterwards
'leaves this answer on the old name. Reading this.sh.Name live would cost a
'crossing on a value that stays the same in practice.
'@return String. The worksheet name.
Public Property Get Name() As String
    Name = this.nm
End Property

'@label:password-sheet
'@prop-title Worksheet hosting the password tables
'@return Worksheet. The password worksheet.
Public Property Get PasswordSheet() As Worksheet
    EnsureSheetReady
    Set PasswordSheet = this.sh
End Property

'@label:display-prompts
'@prop-title Whether UI prompts are displayed during operations
'@return Boolean. True when prompts are enabled.
Public Property Get DisplayPrompts() As Boolean
    DisplayPrompts = this.displayPrompts
End Property

'@label:display-prompts-set
'@prop-title Toggle UI prompt display
'@details
'Set this to False to run the key routines without a dialog. A test harness
'needs it: MsgBox blocks a headless run.
Public Property Let DisplayPrompts(ByVal value As Boolean)
    this.displayPrompts = value
End Property

'@label:last-private-prompt
'@prop-title Most recent private key prompt text
'@return String. The prompt text.
Public Property Get LastPrivatePrompt() As String
    LastPrivatePrompt = this.lastPrivatePrompt
End Property

'@label:last-private-prompt-title
'@prop-title Most recent private key prompt title
'@return String. The prompt title.
Public Property Get LastPrivatePromptTitle() As String
    LastPrivatePromptTitle = this.lastPrivateTitle
End Property


'@label:value
'@sub-title Retrieve a range value
'@details Retrieve the value stored in the named range associated with the supplied tag.
'@param tagName A tag identifier (debuggingpassword, publickey, privatekey, version, debugmode, labpublickey, labprivatekey).
'@return A String, value stored in the named range associated with the tag.
Public Property Get Value(ByVal tagName As String) As String
    Value = TagValue(TagFromName(tagName))
End Property


'@label:tag-value
'@prop-title Retrieve a value from a byte-coded tag
'@param tagId Byte. PasswordTag constant.
'@return String. The stored value.
Private Property Get TagValue(ByVal tagId As Byte) As String
    Dim targetRange As Range
    Set targetRange = NamedRange(TagToRangeName(tagId))
    If Not (targetRange Is Nothing) Then
        TagValue = CStr(targetRange.Value)
    End If
End Property

'@label:table-range
'@prop-title Retrieve the range of a password-managed table
'@details
'Looks up the specified ListObject on the password worksheet and
'returns either the full range or data-only range.
'@param tableName String. Table identifier (T_keys or T_ProtectedSheets). Defaults to TABLEKEYS.
'@param includeHeaders Boolean. When True, includes the header row. Defaults to False.
'@return Range. The requested table range.
Public Property Get TableRange(Optional ByVal tableName As String = TABLEKEYS, _
                               Optional ByVal includeHeaders As Boolean = False) As Range

    Dim lo As ListObject

    Select Case LCase$(Trim$(tableName))
        Case LCase$(TABLEKEYS)
            Set lo = KeysTable()
        Case LCase$(TABLEPROTECTED)
            Set lo = ProtectedSheetsTable()
        Case Else
            ThrowError ProjectError.ElementNotFound, _
                      "Unknown password table '" & tableName & "'"
    End Select

    If includeHeaders Then
        Set TableRange = lo.Range
    Else
        Set TableRange = lo.DataBodyRange
    End If
End Property

'@label:has-checkings
'@prop-title Whether any diagnostic entries have been captured
'@return Boolean. True when at least one entry exists.
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the aggregated checking entries
'@return Checking. The checking object, or Nothing.
Public Property Get CheckingValues() As Checking
    If this.hasChecks Then
        Set CheckingValues = this.checks
    End If
End Property

'@section Debug Mode
'===============================================================================
'Toggle debug mode on the target workbook.

'@label:enter-debug-mode
'@sub-title Put the workbook into a fully unprotected debug state
'@details
'Every worksheet in the target workbook is unprotected. LeaveDebugMode restores
'only the sheets listed in the protection table, so a sheet protected by hand,
'or protected with registerState:=False, comes back unprotected. That is the
'contract of this pair.
'
'The debug-mode flag is always written to the bound worksheet, whichever
'workbook is passed in. Bind the object to the workbook you are toggling.
'@param targetWorkbook Optional Workbook. Target workbook. Defaults to the bound workbook.
'@export
Public Sub EnterDebugMode(Optional ByVal targetWorkbook As Workbook)

    Dim wb As Workbook
    Dim childSheet As Worksheet

    Set wb = ResolveWorkbook(targetWorkbook)

    'Unlock workbook structure before processing child worksheets.
    UnprotectWorkbook wb

    For Each childSheet In wb.Worksheets
        'Each sheet must be explicitly unprotected for unrestricted debugging.
        'The sheet object goes straight through, so nothing looks it up by name again.
        UnprotectSheet childSheet
    Next

    SetDebugModeFlag True
    EnsureProtectedSheetsName wb
    LogInfo "Workbook entered debug mode; protections removed", checkingWarning
End Sub

'@label:leave-debug-mode
'@sub-title Restore protections after debug mode
'@details
'Only the sheets listed in the protection table come back protected. The
'debug-mode flag is always written to the bound worksheet, whichever workbook
'is passed in. Bind the object to the workbook you are toggling.
'@param targetWorkbook Optional Workbook. Target workbook. Defaults to the bound workbook.
'@export
Public Sub LeaveDebugMode(Optional ByVal targetWorkbook As Workbook)

    If Not IsInDebugMode() Then Exit Sub

    Dim wb As Workbook
    Set wb = ResolveWorkbook(targetWorkbook)

    'Reapply worksheet-level protections captured in the metadata table.
    ApplyProtectionMatrix wb
    'Restore workbook structure protection once sheets are locked again.
    ProtectWorkbook wb
    SetDebugModeFlag False
    LogInfo "Workbook protections restored after leaving debug mode", checkingInfo
End Sub


'@label:ensure-protected-sheets-name
'@sub-title Ensure the workbook exposes a named range referencing the protection table
'@details
'The reference is built from this object's own protection table. When the target
'workbook holds no table by that name, the name resolves to #REF! and the
'generated close handler fails on it. CloneToWorksheet avoids that by calling
'through the handler bound to the new sheet.
'@param targetWorkbook Optional Workbook. Target workbook. Defaults to the bound workbook.
'@export
Public Sub EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)

    Dim wb As Workbook
    Dim namedItem As Name
    Dim tableObj As ListObject
    Dim structuredRef As String

    Set wb = ResolveWorkbook(targetWorkbook)
    Set tableObj = ProtectedSheetsTable()

    'Structured references keep the name aligned with the listobject even as rows are added or removed.
    structuredRef = "=" & tableObj.Name & "[#All]"

    On Error Resume Next
        Set namedItem = wb.Names(NAMEPROTECTEDSHEETS)
    On Error GoTo 0

    If namedItem Is Nothing Then
        'Create the workbook-level name relied upon by debug shutdown handlers.
        wb.Names.Add Name:=NAMEPROTECTEDSHEETS, RefersTo:=structuredRef
    ElseIf StrComp(CStr(namedItem.RefersTo), structuredRef, vbBinaryCompare) <> 0 Then
        'Keep the existing name in sync in case the table was cloned or renamed.
        namedItem.RefersTo = structuredRef
    End If
End Sub

'@label:ensure-debug-exit-handler
'@sub-title Inject before-close procedures that leave debug mode automatically
'@param targetWorkbook Optional Workbook. Target workbook. Defaults to the bound workbook.
'@export
Public Sub EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook)

    Dim wb As Workbook
    Dim vbProj As Object
    Dim vbComp As Object

    Set wb = ResolveWorkbook(targetWorkbook)
    EnsureProtectedSheetsName wb

    On Error GoTo NoProject
        Set vbProj = wb.VBProject
    On Error GoTo 0

    If vbProj Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "VBProject is not accessible; enable Trust Access to the VBA project object model"
    End If

    On Error GoTo WorkbookModuleMissing
        Set vbComp = vbProj.VBComponents(wb.CodeName)
    On Error GoTo 0

    'Inject or refresh the guard procedures that revert debug mode automatically.
    InsertOrReplaceProcedure vbComp, "LeaveDebugModeOnClose", BuildLeaveDebugProcedure
    EnsureWorkbookBeforeCloseHook vbComp
    Exit Sub

NoProject:
    ThrowError ProjectError.ObjectNotInitialized, _
              "Unable to access VBProject for workbook '" & wb.Name & "'"
    Exit Sub

WorkbookModuleMissing:
    ThrowError ProjectError.ElementNotFound, _
              "Workbook module '" & wb.CodeName & "' is missing"
End Sub


'@section Protection
'===============================================================================
'Protect and unprotect worksheets and workbooks.

'@label:protect-sheet
'@sub-title Protect a worksheet with the debugging password
'@details
'When registerState is True the two capability flags are folded with the values
'already stored for the sheet, and the result is written back. The fold is
'one-way: once a sheet is stored as "no" for shapes or for row deletion, no
'later call widens it back to "yes".
'
'When registerState is False the incoming flags are used as they are and the
'protection table is left alone.
'@param sheetObject Variant. Worksheet or sheet name.
'@param allowShapes Boolean. Permit shape manipulation. Defaults to True.
'@param allowDeletingRows Boolean. Enable row deletion. Defaults to True.
'@param registerState Boolean. Fold in the stored settings and write the result back. Defaults to True.
Private Sub ProtectSheet(sheetObject As Variant, _
                        Optional ByVal allowShapes As Boolean = True, _
                        Optional ByVal allowDeletingRows As Boolean = True, _
                        Optional ByVal registerState As Boolean = True)

    Dim targetSheet As Worksheet
    Dim allowDraw As Boolean
    Dim allowDelete As Boolean
    Dim storedDraw As Boolean
    Dim storedDelete As Boolean
    Dim storedRow As Long

    EnsureSheetReady

    If IsInDebugMode() Then
        LogInfo "Protection request ignored because workbook is in debug mode", checkingWarning
        Exit Sub
    End If

    Set targetSheet = ResolveSheet(sheetObject)

    allowDraw = allowShapes
    allowDelete = allowDeletingRows

    If registerState Then
        ReadProtectionFlags targetSheet.Name, storedDraw, storedDelete, storedRow
        allowDraw = (allowDraw And storedDraw)
        allowDelete = (allowDelete And storedDelete)

        'The row is written when the sheet has none yet, or when the fold
        'changed a flag. A protect call that changes nothing used to write its
        'three cells anyway, and the show/hide bracket protects on every close:
        'three cell writes per click for a row that already said the same.
        If storedRow = 0 Or allowDraw <> storedDraw Or allowDelete <> storedDelete Then
            PersistProtectionSettings targetSheet.Name, allowDraw, allowDelete, storedRow
        End If
    End If

    ApplySheetProtection targetSheet, allowDraw, allowDelete, DebugPassword()
End Sub

'@label:unprotect-sheet
'@sub-title Unprotect a worksheet using the stored debugging password
'@param sheetObject Optional Variant. Worksheet reference or sheet name. Defaults to the active sheet.
Private Sub UnprotectSheet(Optional sheetObject As Variant)

    Dim targetSheet As Worksheet
    Dim pwd As String

    EnsureSheetReady
    Set targetSheet = ResolveSheet(sheetObject)
    pwd = DebugPassword()

    'Excel raises when the stored password does not open the sheet. That sheet is
    'left as it is and the caller carries on with the rest of its work.
    On Error Resume Next
        If targetSheet.ProtectContents Then targetSheet.Unprotect pwd
    On Error GoTo 0
End Sub

'@label:protect-workbook
'@sub-title Protect workbook structure with the debugging password
'@param targetObject Variant. Workbook reference.
Private Sub ProtectWorkbook(targetObject As Variant)
    Dim pwd As String
    Dim wb As Workbook

    pwd = DebugPassword()
    Set wb = ResolveWorkbook(targetObject)

    wb.Protect Password:=pwd, Structure:=True
End Sub

'@label:unprotect-workbook
'@sub-title Unprotect workbook structure using the debugging password
'@param targetObject Variant. Workbook reference.
Private Sub UnprotectWorkbook(targetObject As Variant)
    Dim pwd As String
    Dim wb As Workbook

    pwd = DebugPassword()
    Set wb = ResolveWorkbook(targetObject)

    'Excel raises when the workbook is already open or the password does not match.
    On Error Resume Next
        wb.Unprotect Password:=pwd
    On Error GoTo 0
End Sub

'@label:resolve-protection
'@sub-title Dispatch protect/unprotect via TypeName
'@details
'Accepts a Workbook, Worksheet, or String. The string "_active"
'targets the active worksheet; "_wbactive" targets the active workbook.
'Nothing shares the workbook branch, where ResolveWorkbook turns it back into
'the bound workbook.
'
'This is the one place the protect/unprotect choice is made. Protect and
'UnProtect both come through here, so the two paths cannot drift apart.
'@param targetObject Optional Variant. Workbook, Worksheet, or String tag. Defaults to "_active".
'@param allowShapes Optional Boolean. Allow shape editing when protected. Defaults to True.
'@param allowDeletingRows Optional Boolean. Allow row deletion when protected. Defaults to True.
'@param registerState Optional Boolean. Persist protection metadata. Defaults to True.
'@param setProtection Optional Boolean. When True, protects; when False, unprotects. Defaults to True.
Private Sub ResolveProtection(Optional targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True, _
                   Optional ByVal setProtection As Boolean = True)

    Dim typeNameValue As String
    Dim targetsWorkbook As Boolean

    typeNameValue = TypeName(targetObject)

    Select Case typeNameValue
    Case "Workbook", "Nothing"
        targetsWorkbook = True
    Case "Worksheet"
        targetsWorkbook = False
    Case "String"
        targetsWorkbook = (targetObject = "_wbactive")
    Case Else
        ThrowError ProjectError.InvalidArgument, _
                      "Expected workbook, worksheet or sheet name, or specific tags" & _
                      "(_active, _wbactive) received " & typeNameValue
    End Select

    If targetsWorkbook Then
        If setProtection Then
            ProtectWorkbook targetObject
        Else
            UnprotectWorkbook targetObject
        End If
        Exit Sub
    End If

    If setProtection Then
        ProtectSheet targetObject, allowShapes, allowDeletingRows, registerState
    Else
        UnprotectSheet targetObject
    End If
End Sub

'@label:protect
'@sub-title Protect a worksheet or workbook
'@param targetObject Optional Variant. Workbook, Worksheet, or String tag. Defaults to "_active".
'@param allowShapes Optional Boolean. Allow shape editing. Defaults to True.
'@param allowDeletingRows Optional Boolean. Allow row deletion. Defaults to True.
'@param registerState Optional Boolean. Persist protection metadata. Defaults to True.
'@export
Public Sub Protect(Optional targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True)

    ResolveProtection targetObject, allowShapes, allowDeletingRows, registerState, setProtection:=True
End Sub

'@label:unprotect
'@sub-title Unprotect a worksheet or workbook
'@param targetObject Optional Variant. Workbook, Worksheet, or String tag. Defaults to "_active".
'@export
Public Sub UnProtect(Optional targetObject As Variant = "_active")
    ResolveProtection targetObject, setProtection:=False
End Sub


'@section Keys
'===============================================================================
'Key generation and display.

'@label:generate-key
'@sub-title Randomly select a public/private key pair
'@param trads TranslationObject. Translation helper for prompts.
'@export
Public Sub GenerateKey(ByVal trads As TranslationObject)

    Dim lo As ListObject
    Dim dataRange As Range
    Dim recordIndex As Long
    Dim publicValue As String
    Dim privateValue As String

    If trads Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation object is required to generate a key"
    End If

    Set lo = KeysTable()
    Set dataRange = lo.DataBodyRange

    If dataRange Is Nothing Then
        LogInfo "Unable to generate key; keys table has no data", checkingWarning
        Exit Sub
    End If

    'Int() gives every row the same chance. CLng() rounds to nearest, which halves
    'the chance of the first and last rows and can overshoot the row count.
    Randomize
    recordIndex = Int(dataRange.Rows.Count * Rnd()) + 1

    publicValue = CStr(dataRange.Cells(recordIndex, 1).Value)
    privateValue = CStr(dataRange.Cells(recordIndex, 2).Value)

    NamedRange(NAMEPUBLICKEY).Value = publicValue
    NamedRange(NAMEPRIVATEKEY).Value = privateValue

    DisplayPrivateKey trads
End Sub

'@label:display-private-key
'@sub-title Display the current private key via a translated prompt
'@details
'The dialog is shown only when DisplayPrompts is True, which is the default.
'The prompt text and title are always recorded, so a caller that suppresses the
'dialog can still read what would have been shown.
'@param trads TranslationObject. Translation helper for the message.
'@export
Public Sub DisplayPrivateKey(ByVal trads As TranslationObject)

    Dim promptText As String
    Dim titleText As String

    If trads Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation object is required to display private key"
    End If

    promptText = trads.TranslatedValue("MSG_Password") & TagValue(PasswordTagPrivate)
    titleText = trads.TranslatedValue("MSG_Title")

    this.lastPrivatePrompt = promptText
    this.lastPrivateTitle = titleText
    LogInfo "Private key prompt prepared: " & promptText, checkingInfo, addTimeStamp:=False

    If this.displayPrompts Then
        MsgBox promptText, vbInformation, titleText
    End If
End Sub

'@section DataExchange
'===============================================================================
'Import, export, and cloning operations.

'@label:export-to-workbook
'@sub-title Export the password worksheet into another workbook
'@param destination Workbook. Target workbook.
'@export
Public Sub ExportToWorkbook(ByVal destination As Workbook)

    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet

    EnsureSheetReady

    If destination Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Destination workbook is required for export"
    End If

    Set sourceSheet = PasswordSheet()

    On Error Resume Next
        Set targetSheet = destination.Worksheets(sourceSheet.Name)
    On Error GoTo 0

    If Not targetSheet Is Nothing Then
        ThrowError ProjectError.ElementShouldNotExists, _
                  "Workbook already contains a worksheet named '" & sourceSheet.Name & "'"
    End If

    '@Ignore AssignmentNotUsed
    Set targetSheet = destination.Worksheets.Add(After:=destination.Worksheets(destination.Worksheets.Count))
    targetSheet.Name = sourceSheet.Name

    CopyListObjects sourceSheet, targetSheet
    CopyNamedRanges sourceSheet, targetSheet
    targetSheet.Visible = xlSheetVeryHidden
End Sub

'@label:import-from
'@sub-title Import keys and metadata from another password handler
'@details
'Three named values move across: the public key, the private key and the
'version. The debugging password, the lab keys and the protection table stay
'as they are on this sheet. The destination keys table is resized to the shape
'of the source table, columns included.
'@param source Passwords. Source password object.
'@export
Public Sub ImportFrom(ByVal source As Passwords)

    Dim loSource As Range
    Dim destinationLo As ListObject
    Dim resizedRange As Range

    If source Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source password object is required for import"
    End If

    NamedRange(NAMEPUBLICKEY).Value = source.Value("publickey")
    NamedRange(NAMEPRIVATEKEY).Value = source.Value("privatekey")
    NamedRange(NAMEVERSION).Value = source.Value("version")

    Set loSource = source.TableRange(TABLEKEYS, includeHeaders:=False)
    Set destinationLo = KeysTable()

    If loSource Is Nothing Then
        'Clear any existing rows so the target mirrors the empty source state.
        If Not destinationLo.DataBodyRange Is Nothing Then destinationLo.DataBodyRange.ClearContents
        LogInfo "Import skipped because source keys table is empty", checkingWarning
        ResetCaches
        Exit Sub
    End If

    Set resizedRange = destinationLo.HeaderRowRange.Resize(loSource.Rows.Count + 1, loSource.Columns.Count)
    destinationLo.Resize resizedRange
    destinationLo.DataBodyRange.Value = loSource.Value

    ResetCaches
End Sub

'@label:new-debuging-password
'@sub-title Update the debugging password stored on the password worksheet
'@details
'Raises ElementNotFound when the named range is gone. Create validates that
'name, so this can only fire when the name was removed after the object was
'built.
'@param newValue String. The new debugging password value.
'@export
Public Sub NewDebugingPassword(ByVal newValue As String)

    EnsureSheetReady

    NamedRange(NAMEDEBUGPASSWORD).Value = newValue

    'The cached copy is now stale.
    this.debugPassword = vbNullString
    this.debugPasswordLoaded = False
End Sub

'@label:clone-to-worksheet
'@sub-title Clone the password worksheet into another sheet
'@param targetSheet Worksheet. Destination worksheet.
'@return Passwords. Handler bound to the cloned sheet.
'@export
Public Function CloneToWorksheet(ByVal targetSheet As Worksheet) As Passwords

    Dim pass As Passwords

    If targetSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target sheet is required when cloning password data"
    End If

    targetSheet.Cells.Clear
    CopyListObjects PasswordSheet(), targetSheet
    CopyNamedRanges PasswordSheet(), targetSheet

    Set pass = Create(targetSheet)
    'The name is written through the new handler, so it points at the new table.
    pass.EnsureProtectedSheetsName

    Set CloneToWorksheet = pass
End Function


'@label:clone-to-workbook
'@sub-title Clone the password worksheet into another workbook
'@details
'Exports the password data to the target workbook and returns a
'new handler bound to the cloned sheet.
'@param wb Workbook. Destination workbook.
'@return Passwords. Handler bound to the cloned sheet.
'@export
Public Function CloneToWorkbook(ByVal wb As Workbook) As Passwords
    Dim sh As Worksheet
    ExportToWorkbook wb
    Set sh = wb.Worksheets(Name())
    Set CloneToWorkbook = Create(sh)
End Function

'@section Private Helpers
'===============================================================================
'Validation, caching, and internal operations.

'@label:validate-sheet
'@sub-title Validate the password worksheet
'@details
'The two lab key names are left out on purpose: password sheets built before
'those names exist still work. CopyNamedRanges skips whichever of them is
'missing and records the gap.
Private Sub ValidateSheet(ByVal sh As Worksheet)
    If sh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Password sheet is required"
    SheetTable sh, TABLEKEYS
    SheetTable sh, TABLEPROTECTED
    NamedRange NAMEDEBUGPASSWORD, sh
    NamedRange NAMEPUBLICKEY, sh
    NamedRange NAMEPRIVATEKEY, sh
    NamedRange NAMEDEBUGMODE, sh
    NamedRange NAMEVERSION, sh
End Sub

'@label:ensure-sheet-ready
'@sub-title Ensure password worksheet is initialised
'@details
'this.wb is written on the same statement as this.sh, and Create is the only
'way in, so a ready sheet always carries a ready workbook.
Private Sub EnsureSheetReady()
    If this.sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Password sheet is not initialised"
    End If
End Sub

'@label:reset-caches
'@sub-title Clear the cached tables, password and debug flag
Private Sub ResetCaches()
    Set this.keyTable = Nothing
    Set this.protectedTable = Nothing
    this.debugPassword = vbNullString
    this.debugPasswordLoaded = False
    this.debugMode = False
    this.debugModeLoaded = False
End Sub

'@label:sheet-table
'@sub-title Resolve a ListObject on a worksheet
'@param targetSheet Worksheet. Sheet holding the table.
'@param tableName String. Name of the table.
'@return ListObject. The resolved table.
Private Function SheetTable(ByVal targetSheet As Worksheet, ByVal tableName As String) As ListObject

    Dim result As ListObject

    On Error Resume Next
        Set result = targetSheet.ListObjects(tableName)
    On Error GoTo 0

    If result Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "ListObject '" & tableName & "' not found on sheet '" & targetSheet.Name & "'"
    End If

    Set SheetTable = result
End Function

'@label:cached-table
'@sub-title Resolve a table on the bound sheet and remember it
'@param tableName String. Name of the table.
'@param cache ListObject. The field that holds the answer between calls.
'@return ListObject. The resolved table.
Private Function CachedTable(ByVal tableName As String, ByRef cache As ListObject) As ListObject
    EnsureSheetReady

    If cache Is Nothing Then Set cache = SheetTable(this.sh, tableName)

    Set CachedTable = cache
End Function

'@label:keys-table
'@sub-title Retrieve and cache the keys ListObject
'@return ListObject. The keys table.
Private Function KeysTable() As ListObject
    Set KeysTable = CachedTable(TABLEKEYS, this.keyTable)
End Function

'@label:protected-sheets-table
'@sub-title Retrieve and cache the protected-sheets ListObject
'@return ListObject. The protection table.
Private Function ProtectedSheetsTable() As ListObject
    Set ProtectedSheetsTable = CachedTable(TABLEPROTECTED, this.protectedTable)
End Function

'@label:named-range
'@sub-title Retrieve a named range
'@details
'Resolves against the bound worksheet, or against targetSheet when one is
'passed. Create validates the sheet before it is bound, which is what the
'second argument is for.
'@param rangeName String. Name to resolve.
'@param targetSheet Optional Worksheet. Sheet to resolve against. Defaults to the bound sheet.
'@return Range. The named range.
Private Function NamedRange(ByVal rangeName As String, _
                            Optional ByVal targetSheet As Worksheet) As Range

    Dim sourceSheet As Worksheet
    Dim result As Range

    If targetSheet Is Nothing Then
        EnsureSheetReady
        Set sourceSheet = this.sh
    Else
        Set sourceSheet = targetSheet
    End If

    On Error Resume Next
        Set result = sourceSheet.Range(rangeName)
    On Error GoTo 0

    If result Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & rangeName & "' is missing on sheet '" & sourceSheet.Name & "'"
    End If

    Set NamedRange = result
End Function

'@label:debug-password
'@sub-title Read the debugging password, once per object
'@details
'The password is one cell that changes at most once in a session, and every
'protect and unprotect call needs it. It is read on the first call and kept.
'ResetCaches and NewDebugingPassword drop the copy.
'@return String. The debugging password.
Private Function DebugPassword() As String
    If Not this.debugPasswordLoaded Then
        this.debugPassword = TagValue(PasswordTagDebug)
        this.debugPasswordLoaded = True
    End If

    DebugPassword = this.debugPassword
End Function

'@label:resolve-workbook
'@sub-title Resolve a workbook from optional candidate
'@return Workbook. The resolved workbook.
Private Function ResolveWorkbook(Optional candidate As Variant) As Workbook

    If IsMissing(candidate) Or IsEmpty(candidate) Then
        EnsureSheetReady
        Set ResolveWorkbook = this.wb
        Exit Function
    End If

    If IsObject(candidate) Then
        If (candidate Is Nothing) Then
            EnsureSheetReady
            Set ResolveWorkbook = this.wb
        ElseIf StrComp(TypeName(candidate), "Workbook", vbTextCompare) = 0 Then
            Set ResolveWorkbook = candidate
        Else
            ThrowError ProjectError.InvalidArgument, _
                      "Expected workbook, or specific tags (_wbactive), received " & TypeName(candidate)
        End If
        Exit Function
    End If

    If VarType(candidate) = vbString Then
        If candidate = "_wbactive" Then
            EnsureSheetReady
            Set ResolveWorkbook = this.wb
        Else
            ThrowError ProjectError.InvalidArgument, _
                      "Expected workbook, or specific tags (_wbactive), received " & TypeName(candidate)
        End If
        Exit Function
    End If

    ThrowError ProjectError.InvalidArgument, _
              "Expected workbook, or specific tags (_wbactive), received " & TypeName(candidate)
End Function

'@label:resolve-sheet
'@sub-title Resolve a worksheet from a Variant argument
'@return Worksheet. The resolved worksheet.
Private Function ResolveSheet(sheetObject As Variant) As Worksheet

    Dim typeNameValue As String
    Dim wb As Workbook
    Dim sh As Worksheet

    typeNameValue = TypeName(sheetObject)

    Select Case typeNameValue
        Case "Worksheet"
            Set ResolveSheet = sheetObject
        Case "String"
            If sheetObject = "_active" Then
                Set ResolveSheet = ActiveSheet
            Else
                Set wb = ResolveWorkbook()
                On Error Resume Next
                    Set sh = wb.Worksheets(sheetObject)
                On Error GoTo 0
                If sh Is Nothing Then
                    ThrowError ProjectError.ElementNotFound, "Worksheet " & sheetObject & " not found"
                End If
                Set ResolveSheet = sh
            End If
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "Expected worksheet or sheet name, received " & typeNameValue
    End Select
End Function

'@label:tag-to-range-name
'@sub-title Map a tag identifier to its named range name
'@return String. The named range identifier.
Private Function TagToRangeName(ByVal tagId As Byte) As String
    Select Case tagId
        Case PasswordTagDebug: TagToRangeName = NAMEDEBUGPASSWORD
        Case PasswordTagPublic: TagToRangeName = NAMEPUBLICKEY
        Case PasswordTagPrivate: TagToRangeName = NAMEPRIVATEKEY
        Case PasswordTagVersion: TagToRangeName = NAMEVERSION
        Case PasswordTagMode: TagToRangeName = NAMEDEBUGMODE
        Case PasswordTagLabPublic: TagToRangeName = NAMELABPUBLICKEY
        Case PasswordTagLabPrivate: TagToRangeName = NAMELABPRIVATEKEY
        Case Else
            ThrowError ProjectError.InvalidArgument, "Unknown password tag identifier: " & CStr(tagId)
    End Select
End Function

'@label:tag-from-name
'@sub-title Convert a friendly tag label to a byte-coded identifier
'@return Byte. The PasswordTag constant.
Private Function TagFromName(ByVal tagName As String) As Byte
    Select Case LCase$(Trim$(tagName))
        Case "debuggingpassword": TagFromName = PasswordTagDebug
        Case "publickey": TagFromName = PasswordTagPublic
        Case "privatekey": TagFromName = PasswordTagPrivate
        Case "version": TagFromName = PasswordTagVersion
        Case "debugmode": TagFromName = PasswordTagMode
        Case "labpublickey": TagFromName = PasswordTagLabPublic
        Case "labprivatekey": TagFromName = PasswordTagLabPrivate
        Case Else
            ThrowError ProjectError.InvalidArgument, "Unknown password tag '" & tagName & "'"
    End Select
End Function

'@label:is-in-debug-mode
'@sub-title Test whether the workbook is in debug mode
'@details
'The flag is one cell on the bound sheet, and ProtectSheet asks for it on every
'call. It is read once and kept. EnterDebugMode and LeaveDebugMode are the two
'writers and both refresh the copy; ResetCaches drops it.
'@return Boolean. True when debug mode is active.
Public Function IsInDebugMode() As Boolean
    If Not this.debugModeLoaded Then
        this.debugMode = (StrComp(CStr(NamedRange(NAMEDEBUGMODE).Value), DEFAULTBOOLYES, vbTextCompare) = 0)
        this.debugModeLoaded = True
    End If

    IsInDebugMode = this.debugMode
End Function

'@label:set-debug-mode-flag
'@sub-title Write the debug-mode flag and refresh the cached copy
'@param active Boolean. True to record debug mode as on.
Private Sub SetDebugModeFlag(ByVal active As Boolean)
    If active Then
        NamedRange(NAMEDEBUGMODE).Value = DEFAULTBOOLYES
    Else
        NamedRange(NAMEDEBUGMODE).Value = DEFAULTBOOLNO
    End If

    this.debugMode = active
    this.debugModeLoaded = True
End Sub

'@label:apply-protection-matrix
'@sub-title Apply stored protection settings to all listed worksheets
'@details
'Reads the whole protection table in one go and protects each listed sheet with
'the same password. A row naming a sheet that was renamed or deleted is written
'to the checking log, and the remaining rows still get their protection.
'@param wb Workbook. Target workbook whose sheets will be protected.
'@export
Public Sub ApplyProtectionMatrix(ByVal wb As Workbook)
    Dim bodyRange As Range
    Dim rowIdx As Long
    Dim sheetName As String
    Dim allowShapesValue As Boolean
    Dim allowDeleteRows As Boolean
    Dim data As Variant
    Dim pwd As String
    Dim protectFailed As Boolean
    Dim failureText As String

    Set bodyRange = ProtectedSheetsTable().DataBodyRange

    If bodyRange Is Nothing Then Exit Sub

    'Read password once (same for all sheets) and bulk-read the table
    pwd = DebugPassword()
    data = bodyRange.Value2

    For rowIdx = 1 To UBound(data, 1)
        sheetName = CStr(data(rowIdx, 1))
        If LenB(sheetName) > 0 Then
            allowShapesValue = ParseBoolean(CStr(data(rowIdx, 2)))
            allowDeleteRows = ParseBoolean(CStr(data(rowIdx, 3)))

            protectFailed = False
            failureText = vbNullString

            On Error Resume Next
                ApplySheetProtection wb.Worksheets(sheetName), allowShapesValue, allowDeleteRows, pwd
                protectFailed = (Err.Number <> 0)
                failureText = Err.Description
                Err.Clear
            On Error GoTo 0

            If protectFailed Then
                LogInfo "Protection was not restored for sheet '" & sheetName & "': " & failureText, _
                        checkingWarning
            End If
        End If
    Next rowIdx
End Sub

'@label:apply-sheet-protection
'@sub-title Protect a worksheet using debugging password and capability flags
'@details
'UserInterfaceOnly keeps macros able to write to the sheet. Excel drops that
'flag when the file closes, so a reopened workbook has to run
'ApplyProtectionMatrix again before code writes anywhere.
Private Sub ApplySheetProtection(ByVal targetSheet As Worksheet, _
                                 ByVal allowShapes As Boolean, _
                                 ByVal allowDeleting As Boolean, _
                                 ByVal pwd As String)

    targetSheet.Protect Password:=pwd, _
        DrawingObjects:=(Not allowShapes), _
        Contents:=True, _
        Scenarios:=True, _
        AllowInsertingRows:=allowDeleting, _
        AllowSorting:=False, _
        AllowFiltering:=True, _
        AllowFormattingCells:=False, _
        AllowDeletingRows:=allowDeleting, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        UserInterfaceOnly:=True
End Sub

'@label:read-protection-flags
'@sub-title Read the stored capability flags for a worksheet
'@details
'The protection table is read as one block and matched in memory, so a lookup
'costs one crossing. Sheet names are matched without regard to case, the way
'Excel itself treats them, so "Data" and "data" resolve to the same row.
'@param sheetName String. Sheet to look up.
'@param allowShapes Boolean. Receives the stored shape flag, True when the sheet has no row.
'@param allowDeletingRows Boolean. Receives the stored row-deletion flag, True when the sheet has no row.
'@param rowIndex Long. Receives the 1-based data row, 0 when the sheet has no row.
Private Sub ReadProtectionFlags(ByVal sheetName As String, _
                                ByRef allowShapes As Boolean, _
                                ByRef allowDeletingRows As Boolean, _
                                ByRef rowIndex As Long)

    Dim bodyRange As Range
    Dim data As Variant
    Dim rowIdx As Long

    'Defaults for a sheet the table has never seen.
    allowShapes = True
    allowDeletingRows = True
    rowIndex = 0

    Set bodyRange = ProtectedSheetsTable().DataBodyRange
    If bodyRange Is Nothing Then Exit Sub

    data = bodyRange.Value2
    If Not IsArray(data) Then Exit Sub
    If UBound(data, 2) < PROTECTIONCOLUMNS Then Exit Sub

    For rowIdx = 1 To UBound(data, 1)
        If StrComp(CStr(data(rowIdx, 1)), sheetName, vbTextCompare) = 0 Then
            allowShapes = ParseBoolean(CStr(data(rowIdx, 2)))
            allowDeletingRows = ParseBoolean(CStr(data(rowIdx, 3)))
            rowIndex = rowIdx
            Exit Sub
        End If
    Next rowIdx
End Sub

'@label:persist-protection-settings
'@sub-title Insert or update the protection table entry for a sheet
'@param sheetName String. Sheet the row describes.
'@param allowShapes Boolean. Shape flag to store.
'@param allowDeletingRows Boolean. Row-deletion flag to store.
'@param rowIndex Long. 1-based data row to overwrite. 0 appends a new row.
Private Sub PersistProtectionSettings(ByVal sheetName As String, _
                                      ByVal allowShapes As Boolean, _
                                      ByVal allowDeletingRows As Boolean, _
                                      ByVal rowIndex As Long)

    Dim lo As ListObject
    Dim targetRow As Range

    Set lo = ProtectedSheetsTable()

    If rowIndex > 0 Then
        Set targetRow = lo.DataBodyRange.Rows(rowIndex)
    Else
        'Append a new row when the sheet was not found in the table.
        Set targetRow = lo.ListRows.Add.Range
    End If

    targetRow.Cells(1, 1).Value = sheetName
    targetRow.Cells(1, 2).Value = BooleanToFlag(allowShapes)
    targetRow.Cells(1, 3).Value = BooleanToFlag(allowDeletingRows)
End Sub


'@label:boolean-to-flag
'@sub-title Convert a Boolean to "yes"/"no" text flag
'@return String. "yes" or "no".
Private Function BooleanToFlag(ByVal inpValue As Boolean) As String
    If inpValue Then
        BooleanToFlag = DEFAULTBOOLYES
    Else
        BooleanToFlag = DEFAULTBOOLNO
    End If
End Function

'@label:parse-boolean
'@sub-title Parse a "yes"/"no" text flag to Boolean
'@return Boolean. True unless the value is "no".
Private Function ParseBoolean(ByVal inpValue As String) As Boolean
    ParseBoolean = (StrComp(CStr(inpValue), DEFAULTBOOLNO, vbTextCompare) <> 0)
End Function


'@label:build-leave-debug-procedure
'@sub-title Generate VBA code for the LeaveDebugModeOnClose procedure
'@return String. VBA source code.
Private Function BuildLeaveDebugProcedure() As String
    Dim code As String
    code = "Private Sub LeaveDebugModeOnClose()" & vbNewLine & _
           "    On Error Resume Next" & vbNewLine & _
           "    Dim guardName As Name" & vbNewLine & _
           "    Set guardName = ThisWorkbook.Names(""" & NAMEPROTECTEDSHEETS & """)" & vbNewLine & _
           "    If guardName Is Nothing Then Exit Sub" & vbNewLine & _
           "    Dim targetSheet As Worksheet" & vbNewLine & _
           "    Set targetSheet = guardName.RefersToRange.Worksheet" & vbNewLine & _
           "    Dim passObj As Passwords" & vbNewLine & _
           "    Set passObj = Passwords.Create(targetSheet)" & vbNewLine & _
           "    passObj.LeaveDebugMode ThisWorkbook" & vbNewLine & _
           "End Sub"
    BuildLeaveDebugProcedure = code
End Function

'@label:build-before-close-procedure
'@sub-title Generate VBA code for the Workbook_BeforeClose handler
'@return String. VBA source code.
Private Function BuildBeforeCloseProcedure() As String
    Dim code As String
    code = "Private Sub Workbook_BeforeClose(Cancel As Boolean)" & vbNewLine & _
           "    On Error Resume Next" & vbNewLine & _
           "    If Not Cancel Then LeaveDebugModeOnClose" & vbNewLine & _
           "End Sub"
    BuildBeforeCloseProcedure = code
End Function

'@label:ensure-workbook-before-close-hook
'@sub-title Ensure Workbook_BeforeClose invokes the debug guard
Private Sub EnsureWorkbookBeforeCloseHook(ByVal component As Object)
    Dim codeModule As Object
    Dim procedureName As String
    Dim procStart As Long
    Dim procLines As Long
    Dim bodyStart As Long
    Dim guardToken As String

    procedureName = "Workbook_BeforeClose"
    guardToken = "LeaveDebugModeOnClose"

    Set codeModule = component.CodeModule
    If codeModule Is Nothing Then
        ThrowError ProjectError.SomethingWentWrong, _
                   "Unable to access workbook module when installing Workbook_BeforeClose guard"
    End If

    On Error Resume Next
    procStart = codeModule.ProcStartLine(procedureName, 0)
    On Error GoTo Handler

    If procStart <= 0 Then
        'Procedure missing: append a fresh stub without modifying the existing module surface.
        AppendProcedure codeModule, BuildBeforeCloseProcedure
        GoTo Success
    End If

    procLines = codeModule.ProcCountLines(procedureName, 0)

    If ProcedureBlockContainsToken(codeModule, procStart, procLines, guardToken) Then GoTo Success

    bodyStart = codeModule.ProcBodyLine(procedureName, 0)
    codeModule.InsertLines bodyStart + 1, "On Error Resume Next"
    codeModule.InsertLines bodyStart + 2, "    If Not Cancel Then LeaveDebugModeOnClose"
Success:
    Exit Sub

Handler:
    ThrowError ProjectError.SomethingWentWrong, _
              "Unable to ensure Workbook_BeforeClose invokes LeaveDebugModeOnClose: " & Err.Description
End Sub

'@label:procedure-block-contains-token
'@sub-title Scan a procedure block for a specific token
'@return Boolean. True when the token is found.
Private Function ProcedureBlockContainsToken(ByVal codeModule As Object, _
                                             ByVal startLine As Long, _
                                             ByVal lineCount As Long, _
                                             ByVal token As String) As Boolean
    Dim blockText As String

    If startLine <= 0 Or lineCount <= 0 Then Exit Function

    blockText = codeModule.Lines(startLine, lineCount)
    ProcedureBlockContainsToken = (InStr(1, blockText, token, vbTextCompare) > 0)
End Function

'@label:append-procedure
'@sub-title Append generated code to the end of a code module
Private Sub AppendProcedure(ByVal codeModule As Object, ByVal procedureCode As String)
    Dim lineCount As Long

    lineCount = codeModule.CountOfLines
    If lineCount > 0 Then lineCount = lineCount + 1
    codeModule.InsertLines lineCount + 1, procedureCode
End Sub


'@label:insert-or-replace-procedure
'@sub-title Insert or replace a VBA procedure in a code module
Private Sub InsertOrReplaceProcedure(ByVal component As Object, _
                                     ByVal procedureName As String, _
                                     ByVal procedureCode As String)

    Dim codeModule As Object
    Dim startLine As Long
    Dim endLine As Long

    Set codeModule = component.CodeModule

    If codeModule Is Nothing Then
        ThrowError ProjectError.SomethingWentWrong, "Unable to access code module for VB component"
    End If

    On Error Resume Next
        startLine = codeModule.ProcStartLine(procedureName, 0)
        If startLine > 0 Then
            endLine = codeModule.ProcCountLines(procedureName, 0)
            codeModule.DeleteLines startLine, endLine
        End If
    On Error GoTo 0

    'Append the generated procedure at the end of the module.
    AppendProcedure codeModule, procedureCode
End Sub

'@label:copy-list-objects
'@sub-title Replicate ListObjects from source to target sheet
Private Sub CopyListObjects(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)

    Dim lo As ListObject
    Dim targetRange As Range

    For Each lo In sourceSheet.ListObjects
        Set targetRange = targetSheet.Range(lo.Range.Address)
        'Copy values first so table resize includes data.
        targetRange.Value = lo.Range.Value
        With targetSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
            .Name = lo.Name
            .TableStyle = lo.TableStyle
            .HeaderRowRange.Interior.Color = lo.HeaderRowRange.Interior.Color
            .HeaderRowRange.Font.Color = lo.HeaderRowRange.Font.Color
        End With
    Next lo
End Sub

'@label:copy-named-ranges
'@sub-title Copy critical named ranges to the target sheet
'@details
'The two lab key names are younger than the other five, so a password sheet
'built before them passes ValidateSheet without them. Each missing name is
'skipped and written to the checking log, which keeps an old password sheet
'usable at build time.
Private Sub CopyNamedRanges(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)

    Dim rangeNames As Variant
    Dim idx As Long
    Dim rangeName As String
    Dim sourceRange As Range
    Dim targetRange As Range

    rangeNames = Array(NAMEDEBUGPASSWORD, NAMEPUBLICKEY, NAMELABPUBLICKEY, _
                       NAMEPRIVATEKEY, NAMELABPRIVATEKEY, NAMEDEBUGMODE, NAMEVERSION)

    For idx = LBound(rangeNames) To UBound(rangeNames)
        rangeName = CStr(rangeNames(idx))
        Set sourceRange = Nothing

        On Error Resume Next
            Set sourceRange = sourceSheet.Range(rangeName)
        On Error GoTo 0

        If sourceRange Is Nothing Then
            LogInfo "Named range '" & rangeName & "' is missing on sheet '" & _
                    sourceSheet.Name & "' and was not copied", checkingWarning
        Else
            Set targetRange = targetSheet.Range(sourceRange.Address)
            targetRange.Value = sourceRange.Value
            If sourceRange.HasFormula Then targetRange.Formula = sourceRange.Formula
            On Error Resume Next
                targetSheet.Names(rangeName).Delete
            On Error GoTo 0
            'Recreate the name so the clone is self-contained.
            targetSheet.Names.Add Name:=rangeName, RefersTo:=targetRange, Visible:=False
        End If
    Next idx
End Sub

'@label:log-info
'@sub-title Append a diagnostic entry to the internal checking log
'@param label String. Message describing the event.
'@param scope Byte. Checking severity. Defaults to checkingNote.
'@param addTimeStamp Boolean. Whether to prepend a timestamp. Defaults to True.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote, _
                    Optional ByVal addTimeStamp As Boolean = True)

    Dim descrip As String

    If this.checks Is Nothing Then
        Set this.checks = Checking.Create(CHECKINGCONTEXT)
    End If

    descrip = label
    If addTimeStamp Then descrip = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - " & descrip

    this.checks.Add CStr(this.checkCounter), descrip, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:throw-error
'@sub-title Raise a typed project error
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub
