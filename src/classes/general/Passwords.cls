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
'@depends Checking, BetterArray
'@IgnoreModule ProcedureNotUsed, ReadOnlyPropertyAssignment, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IPasswords

'@class Passwords
'@description
'Manages workbook/worksheet protection, key generation, debug-mode
'toggling, and import/export of password metadata. Uses TypeName
'dispatch to handle Workbook, Worksheet, and String target arguments
'through a single Protect/UnProtect entry point.
'@depends Checking, BetterArray, ProjectError, ITranslationObject
'@version 1.0 (2026-02-09)
'@author OutbreakTools

'@section State and Constants
'===============================================================================
Private Type TPasswords
    sh As Worksheet                          'Password worksheet backing the object
    wb As Workbook                           'Cached workbook reference
    keyTable As ListObject                   'Cached list object holding keys
    protectedTable As ListObject             'Cached list object holding protection metadata
    hasChecks As Boolean                     'Whether checkings were emitted
    checks As IChecking                      'Checking aggregation object
    checkCounter As Long                     'Incremental counter for checking keys
    nm As String                             'Name of the password worksheet
    displayPrompts As Boolean                'Whether UI prompts should be displayed
    lastPrivatePrompt As String              'Most recent private key prompt text
    lastPrivateTitle As String               'Most recent private key prompt title
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



'@section Core Accessors
'===============================================================================
'Factory, accessors, and stored-value retrieval.

'@label:create
'@sub-title Create an instance of password
'@details
'Factory returning an interface reference that wraps the provided
'password worksheet.
'@param sh Worksheet. Worksheet hosting password tables and named ranges.
'@return IPasswords. Fully initialised password handler.
'@export
Public Function Create(ByVal sh As Worksheet) As IPasswords

    ValidateSheet sh

    With New Passwords
        Set .PasswordSheet = sh
        Set Create = .Self
    End With

End Function

'@label:self
'@prop-title Self-reference as IPasswords
'@return IPasswords. Self-reference.
Public Property Get Self() As IPasswords
    Set Self = Me
End Property

'@label:password-sheet-set
'@prop-title Assign the password worksheet and reset caches
'@param sh Worksheet. The password worksheet to bind.
Public Property Set PasswordSheet(ByVal sh As Worksheet)
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

'@label:password-workbook
'@prop-title Workbook owning the password worksheet
'@return Workbook. The parent workbook.
Public Property Get PasswordWorkbook() As Workbook
    EnsureWorkbookReady
    Set PasswordWorkbook = this.wb
End Property

'@label:display-prompts
'@prop-title Whether UI prompts are displayed during operations
'@return Boolean. True when prompts are enabled.
Public Property Get DisplayPrompts() As Boolean
    DisplayPrompts = this.displayPrompts
End Property

'@label:display-prompts-set
'@prop-title Toggle UI prompt display
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
'@param tagName A tag identifier (debuggingpassword, publickey, privatekey, version, labpublickey, labprivatekey).
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

    Dim Lo As ListObject

    Select Case LCase$(Trim$(tableName))
        Case LCase$(TABLEKEYS)
            Set Lo = KeysTable()
        Case LCase$(TABLEPROTECTED)
            Set Lo = ProtectedSheetsTable()
        Case Else
            ThrowError ProjectError.ElementNotFound, _
                      "Unknown password table '" & tableName & "'"
    End Select
 
    If includeHeaders Then
        Set TableRange = Lo.Range
    Else
        Set TableRange = Lo.DataBodyRange
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
'@return IChecking. The checking object, or Nothing.
Public Property Get CheckingValues() As IChecking
    If this.hasChecks Then
        Set CheckingValues = this.checks
    End If
End Property

'@section Debug Mode
'===============================================================================
'Toggle debug mode on the target workbook.

'@label:enter-debug-mode
'@sub-title Put the workbook into a fully unprotected debug state
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
        UnprotectSheet childSheet.Name
    Next

    NamedRange(NAMEDEBUGMODE).Value = DEFAULTBOOLYES
    EnsureProtectedSheetsName wb
    LogInfo "Workbook entered debug mode; protections removed", checkingWarning
End Sub

'@label:leave-debug-mode
'@sub-title Restore protections after debug mode
'@param targetWorkbook Optional Workbook. Target workbook. Defaults to the bound workbook.
'@export
Public Sub LeaveDebugMode(Optional ByVal targetWorkbook As Workbook)

    If Not InDebugMode() Then Exit Sub

    Dim wb As Workbook
    Set wb = ResolveWorkbook(targetWorkbook)

    'Reapply worksheet-level protections captured in the metadata table.
    ApplyProtectionMatrix wb
    'Restore workbook structure protection once sheets are locked again.
    ProtectWorkbook wb
    NamedRange(NAMEDEBUGMODE).Value = DEFAULTBOOLNO
    LogInfo "Workbook protections restored after leaving debug mode", checkingInfo
End Sub


'@label:ensure-protected-sheets-name
'@sub-title Ensure the workbook exposes a named range referencing the protection table
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
'@param sheetObject Variant. Worksheet or sheet name.
'@param allowShapes Boolean. Permit shape manipulation. Defaults to True.
'@param allowDeletingRows Boolean. Enable row deletion. Defaults to True.
'@param registerState Boolean. Reuse stored table settings. Defaults to True.
Private Sub ProtectSheet(sheetObject As Variant, _
                        Optional ByVal allowShapes As Boolean = True, _
                        Optional ByVal allowDeletingRows As Boolean = True, _
                        Optional ByVal registerState As Boolean = True)

    Dim targetSheet As Worksheet
    Dim allowDraw As Boolean
    Dim allowDelete As Boolean
    Dim record As BetterArray

    EnsureWorkbookReady

    If InDebugMode() Then
        LogInfo "Protection request ignored because workbook is in debug mode", checkingWarning
        Exit Sub
    End If

    Set targetSheet = ResolveSheet(sheetObject)

    If registerState Then
        'Reuse existing stored preferences to guarantee consistent application.
        Set record = ProtectionRecord(targetSheet.Name)
        allowDraw = ParseBoolean(record.Item(record.LowerBound + 1))
        allowDelete = ParseBoolean(record.Item(record.LowerBound + 2))
    End If

    'If one of the two (registered or incoming is false, Set to false)
    allowDraw = (allowShapes And allowDraw)
    allowDelete = (allowDeletingRows And AllowDelete)


    If registerState Then PersistProtectionSettings targetSheet.Name, allowDraw, allowDelete
    ApplySheetProtection targetSheet, allowDraw, allowDelete
End Sub

'@label:unprotect-sheet
'@sub-title Unprotect a worksheet using the stored debugging password
'@param sheetObject Optional Variant. Worksheet reference or sheet name. Defaults to the active sheet.
Private Sub UnprotectSheet(Optional sheetObject As Variant)

    Dim targetSheet As Worksheet
    Dim pwd As String

    EnsureWorkbookReady
    Set targetSheet = ResolveSheet(sheetObject)
    pwd = TagValue(PasswordTagDebug)

    On Error Resume Next
        If targetSheet.ProtectContents Then targetSheet.Unprotect pwd
    On Error GoTo 0
End Sub

'@label:resolve-worksheet-protection
'@sub-title Dispatch worksheet protect or unprotect
'@param sheetObject Variant. Worksheet reference or sheet name.
'@param allowShapes Optional Boolean. Allow shape editing when protected. Defaults to True.
'@param allowDeletingRows Optional Boolean. Allow row deletion when protected. Defaults to True.
'@param registerState Optional Boolean. Persist protection metadata. Defaults to True.
'@param setProtection Optional Boolean. When True, protects; when False, unprotects. Defaults to True.
Private Sub ResolveWorksheetProtection( sheetObject As Variant, _
                        Optional ByVal allowShapes As Boolean = True, _
                        Optional ByVal allowDeletingRows As Boolean = True, _
                        Optional ByVal registerState As Boolean = True, _
                        Optional ByVal setProtection As Boolean = True)

    If setProtection Then
        ProtectSheet sheetObject, allowShapes, allowDeletingRows, registerState
    Else 
        UnprotectSheet sheetObject
    End If
End Sub

'@label:protect-workbook
'@sub-title Protect workbook structure with the debugging password
'@param targetObject Variant. Workbook reference.
Private Sub ProtectWorkbook(targetObject As Variant)
    Dim pwd As String
    Dim wb As Workbook

    pwd = TagValue(PasswordTagDebug)
    Set wb = ResolveWorkbook(targetObject)

    wb.Protect Password:=pwd, Structure:=True
End Sub

'@label:unprotect-workbook
'@sub-title Unprotect workbook structure using the debugging password
'@param targetObject Variant. Workbook reference.
Private Sub UnprotectWorkbook(targetObject As Variant)
    Dim pwd As String
    Dim wb As Workbook

    pwd = TagValue(PasswordTagDebug)
    Set wb = ResolveWorkbook(targetObject)

    On Error Resume Next
        wb.Unprotect Password:=pwd
    On Error GoTo 0
End Sub

'@label:resolve-workbook-protection
'@sub-title Dispatch workbook protect or unprotect
'@param targetObject Variant. Workbook reference.
'@param setProtection Optional Boolean. When True, protects; when False, unprotects. Defaults to True.
Private Sub ResolveWorkbookProtection(targetObject As Variant, Optional ByVal setProtection As Boolean = True)
    If setProtection Then
        ProtectWorkbook targetObject
    Else
        UnprotectWorkbook targetObject
    End If
End Sub

'@label:resolve-protection
'@sub-title Dispatch protect/unprotect via TypeName
'@details
'Accepts a Workbook, Worksheet, or String. The string "_active"
'targets the active worksheet; "_wbactive" targets the active workbook.
'@param targetObject Optional Variant. Workbook, Worksheet, or String tag. Defaults to "_active".
'@param allowShapes Optional Boolean. Allow shape editing when protected. Defaults to True.
'@param allowDeletingRows Optional Boolean. Allow row deletion when protected. Defaults to True.
'@param registerState Optional Boolean. Persist protection metadata. Defaults to True.
'@param setProtection Optional Boolean. When True, protects; when False, unprotects. Defaults to True.
Private Sub ResolveProtection(Optional  targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True, _
                   Optional ByVal setProtection As Boolean = True)

    Dim typeNameValue As String

    typeNameValue = TypeName(targetObject)

    Select Case typeNameValue

    Case "Workbook", "Nothing"
        ResolveWorkbookProtection targetObject, setProtection
    Case "Worksheet"
        ResolveWorksheetProtection  targetObject, allowShapes, allowDeletingRows, registerState
    Case "String"
        If (targetObject = "_wbactive") Then
            ResolveWorkbookProtection targetObject, setProtection
        Else
            ResolveWorksheetProtection  targetObject, allowShapes, allowDeletingRows, registerState, setProtection
        End If
    Case Else
        ThrowError ProjectError.InvalidArgument, _
                      "Expected workbook, worksheet or sheet name, or specific tags" & _
                      "(_active, _wbactive) received " & typeNameValue
    End Select
End Sub

'@label:protect
'@sub-title Protect a worksheet or workbook
'@param targetObject Optional Variant. Workbook, Worksheet, or String tag. Defaults to "_active".
'@param allowShapes Optional Boolean. Allow shape editing. Defaults to True.
'@param allowDeletingRows Optional Boolean. Allow row deletion. Defaults to True.
'@param registerState Optional Boolean. Persist protection metadata. Defaults to True.
'@export
Private Sub Protect(Optional  targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True)

    ResolveProtection targetObject, allowShapes, allowDeletingRows, registerState, setProtection:=True
End Sub

'@label:unprotect
'@sub-title Unprotect a worksheet or workbook
'@param targetObject Optional Variant. Workbook, Worksheet, or String tag. Defaults to "_active".
'@export
Private Sub UnProtect(Optional targetObject As Variant = "_active")
    ResolveProtection targetObject, setProtection:=False
End Sub


'@section Keys
'===============================================================================
'Key generation and display.

'@label:generate-key
'@sub-title Randomly select a public/private key pair
'@param trads ITranslationObject. Translation helper for prompts.
'@export
Private Sub GenerateKey(ByVal trads As ITranslationObject)

    Dim Lo As ListObject
    Dim dataRange As Range
    Dim recordIndex As Long
    Dim publicValue As String
    Dim privateValue As String

    If trads Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation object is required to generate a key"
    End If

    Set Lo = KeysTable()
    Set dataRange = Lo.DataBodyRange

    If dataRange Is Nothing Then
        LogInfo "Unable to generate key; keys table has no data", checkingWarning
        Exit Sub
    End If

    Randomize
    recordIndex = CLng((dataRange.Rows.Count * Rnd()) + 1)
    If recordIndex > dataRange.Rows.Count Then recordIndex = dataRange.Rows.Count
    If recordIndex < 1 Then recordIndex = 1

    publicValue = CStr(dataRange.Cells(recordIndex, 1).Value)
    privateValue = CStr(dataRange.Cells(recordIndex, 2).Value)

    NamedRange(NAMEPUBLICKEY).Value = publicValue
    NamedRange(NAMEPRIVATEKEY).Value = privateValue

    DisplayPrivateKey trads
End Sub

'@label:display-private-key
'@sub-title Display the current private key via a translated prompt
'@param trads ITranslationObject. Translation helper for the message.
'@export
Private Sub DisplayPrivateKey(ByVal trads As ITranslationObject)

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
Private Sub ExportToWorkbook(ByVal destination As Workbook)

    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet

    EnsureWorkbookReady

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
'@param source IPasswords. Source password object.
'@export
Private Sub ImportFrom(ByVal source As IPasswords)

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
'@param newValue String. The new debugging password value.
'@export
Private Sub UpdateDebugingPassword(ByVal newValue As String)
    Dim target As Range

    EnsureSheetReady

    On Error Resume Next
        Set target = NamedRange(NAMEDEBUGPASSWORD)
    On Error GoTo 0

    If target Is Nothing Then Exit Sub

    target.Value = newValue
End Sub

'@label:clone-to-worksheet
'@sub-title Clone the password worksheet into another sheet
'@param targetSheet Worksheet. Destination worksheet.
'@return IPasswords. Handler bound to the cloned sheet.
'@export
Private Function CloneToWorksheet(ByVal targetSheet As Worksheet) As IPasswords

    Dim pass As IPasswords

    If targetSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target sheet is required when cloning password data"
    End If

    targetSheet.Cells.Clear
    CopyListObjects PasswordSheet(), targetSheet
    CopyNamedRanges PasswordSheet(), targetSheet

    Set pass = Create(targetSheet)
    pass.EnsureProtectedSheetsName

    Set CloneToWorksheet = pass
End Function


'@label:clone-to-workbook
'@sub-title Clone the password worksheet into another workbook
'@details
'Exports the password data to the target workbook and returns a
'new handler bound to the cloned sheet.
'@param wb Workbook. Destination workbook.
'@return IPasswords. Handler bound to the cloned sheet.
'@export
Private Function CloneToWorkbook(ByVal wb As Workbook) As IPasswords
    Dim sh As Worksheet
    ExportToWorkbook wb
    Set sh = wb.Worksheets(Name())
    Set CloneToWorkbook = Create(sh)
End Function


'@section Interface
'===============================================================================
'IPasswords delegation wiring.

Private Property Get IPasswords_PasswordSheet() As Worksheet
    Set IPasswords_PasswordSheet = PasswordSheet()
End Property

Private Property Get IPasswords_Value(ByVal tagName As String) As String
    IPasswords_Value = Value(tagName)
End Property

Private Property Get IPasswords_DisplayPrompts() As Boolean
    IPasswords_DisplayPrompts = DisplayPrompts
End Property

Private Property Let IPasswords_DisplayPrompts(ByVal value As Boolean)
    DisplayPrompts = value
End Property

Private Property Get IPasswords_LastPrivatePrompt() As String
    IPasswords_LastPrivatePrompt = LastPrivatePrompt
End Property

Private Property Get IPasswords_LastPrivatePromptTitle() As String
    IPasswords_LastPrivatePromptTitle = LastPrivatePromptTitle
End Property

Private Property Get IPasswords_TableRange(ByVal tableName As String, Optional ByVal includeHeaders As Boolean = False) As Range
    Set IPasswords_TableRange = TableRange(tableName, includeHeaders:=includeHeaders)
End Property

Private Property Get IPasswords_HasCheckings() As Boolean
    IPasswords_HasCheckings = HasCheckings
End Property

Private Property Get IPasswords_IsInDebugMode() As Boolean
    IPasswords_IsInDebugMode = InDebugMode()
End Property

Private Property Get IPasswords_CheckingValues() As IChecking
    Set IPasswords_CheckingValues = CheckingValues
End Property

Private Sub IPasswords_EnterDebugMode(Optional ByVal targetWorkbook As Workbook)
    EnterDebugMode targetWorkbook
End Sub

Private Sub IPasswords_LeaveDebugMode(Optional ByVal targetWorkbook As Workbook)
    LeaveDebugMode targetWorkbook
End Sub

Private Sub IPasswords_EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook)
    EnsureDebugExitHandler targetWorkbook
End Sub

Private Sub IPasswords_Protect(Optional targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True)
    Protect targetObject, allowShapes, allowDeletingRows, registerState
End Sub

Private Sub IPasswords_UnProtect(Optional targetObject As Variant = "_active")
    UnProtect targetObject
End Sub

Private Sub IPasswords_GenerateKey(ByVal trads As ITranslationObject)
    GenerateKey trads
End Sub

Private Sub IPasswords_DisplayPrivateKey(ByVal trads As ITranslationObject)
    DisplayPrivateKey trads
End Sub

Private Sub IPasswords_ExportToWorkbook(ByVal wb As Workbook)
    ExportToWorkbook wb
End Sub

Private Sub IPasswords_ImportFrom(ByVal source As IPasswords)
    ImportFrom source
End Sub

Private Sub IPasswords_NewDebugingPassword(ByVal newValue As String)
    UpdateDebugingPassword newValue
End Sub

Private Function IPasswords_CloneToWorksheet(ByVal targetSheet As Worksheet) As IPasswords
    Set IPasswords_CloneToWorksheet = CloneToWorksheet(targetSheet)
End Function

Private Function IPasswords_CloneToWorkbook(ByVal wb As Workbook) As IPasswords
    Set IPasswords_CloneToWorkbook = CloneToWorkbook(wb)
End Function

Private Sub IPasswords_EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)
    EnsureProtectedSheetsName targetWorkbook
End Sub

Private Sub IPasswords_ApplyProtectionMatrix(ByVal wb As Workbook)
    ApplyProtectionMatrix wb
End Sub

'@section Private Helpers
'===============================================================================
'Validation, caching, and internal operations.

'@label:validate-sheet
'@sub-title Validate the password worksheet
Private Sub ValidateSheet(ByVal sh As Worksheet)
    If sh Is Nothing Then ThrowError ProjectError.ObjectNotInitialized, "Password sheet is required"
    EnsureLoExists sh, TABLEKEYS
    EnsureLoExists sh,TABLEPROTECTED
    EnsureRangeExists sh, NAMEDEBUGPASSWORD
    EnsureRangeExists sh, NAMEPUBLICKEY
    EnsureRangeExists sh, NAMEPRIVATEKEY
    EnsureRangeExists sh, NAMEDEBUGMODE
    EnsureRangeExists sh, NAMEVERSION
End Sub

'@label:ensure-sheet-ready
'@sub-title Ensure password worksheet is initialised
Private Sub EnsureSheetReady()
    If this.sh Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Password sheet is not initialised"
    End If
End Sub

'@label:ensure-workbook-ready
'@sub-title Ensure the password workbook reference is resolved
Private Sub EnsureWorkbookReady()
    EnsureSheetReady
    If this.wb Is Nothing Then Set this.wb = this.sh.Parent
End Sub

'@label:ensure-lo-exists
'@sub-title Validate that a ListObject exists on a worksheet
Private Sub EnsureLoExists(ByVal sh As Worksheet, ByVal loName As String)
    Dim Lo As ListObject

    On Error Resume Next
    Set Lo = sh.ListObjects(loName)
    On Error GoTo 0

    If Lo Is Nothing Then
        ThrowError ProjectError.ElementNotFound, loName & " does not exists in worksheet " & sh.Name
    End If
End Sub

'@label:reset-caches
'@sub-title Clear cached ListObject references
Private Sub ResetCaches()
    Set this.keyTable = Nothing 'ListObject
    Set this.protectedTable = Nothing 'ListObject
End Sub

'@label:keys-table
'@sub-title Retrieve and cache the keys ListObject
'@return ListObject. The keys table.
Private Function KeysTable() As ListObject
    EnsureSheetReady
    If this.keyTable Is Nothing Then
        Set this.keyTable = this.sh.ListObjects(TABLEKEYS)
        If this.keyTable Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                      "ListObject '" & TABLEKEYS & "' not found on sheet '" & this.sh.Name & "'"
        End If
    End If
    Set KeysTable = this.keyTable
End Function

'@label:protected-sheets-table
'@sub-title Retrieve and cache the protected-sheets ListObject
'@return ListObject. The protection table.
Private Function ProtectedSheetsTable() As ListObject
    EnsureSheetReady
    If this.protectedTable Is Nothing Then
        Set this.protectedTable = this.sh.ListObjects(TABLEPROTECTED)
        If this.protectedTable Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                      "ListObject '" &TABLEPROTECTED & "' not found on sheet '" & this.sh.Name & "'"
        End If
    End If
    Set ProtectedSheetsTable = this.protectedTable
End Function

'@label:named-range
'@sub-title Retrieve a named range from the password worksheet
'@return Range. The named range.
Private Function NamedRange(ByVal rangeName As String) As Range

    Dim result As Range

    EnsureSheetReady

    On Error Resume Next
        Set result = this.sh.Range(rangeName)
    On Error GoTo 0

    If result Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & rangeName & "' is missing on sheet '" & this.sh.Name & "'"
    End If

    Set NamedRange = result
End Function

'@label:resolve-workbook
'@sub-title Resolve a workbook from optional candidate
'@return Workbook. The resolved workbook.
Private Function ResolveWorkbook(Optional candidate As Variant) As Workbook

    If IsMissing(candidate) Or IsEmpty(candidate) Then
        EnsureWorkbookReady
        Set ResolveWorkbook = this.wb
        Exit Function
    End If

    If IsObject(candidate) Then
        If (candidate Is Nothing) Then
            EnsureWorkbookReady
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
            EnsureWorkbookReady
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
            If  sheetObject =  "_active" Then
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

'@label:ensure-range-exists
'@sub-title Validate that a named range exists on a worksheet
Private Sub EnsureRangeExists(ByVal targetSheet As Worksheet, ByVal rangeName As String)
    Dim rng As Range
    On Error Resume Next
        Set rng = targetSheet.Range(rangeName)
    On Error GoTo 0
    If rng Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & rangeName & "' is missing on sheet '" & targetSheet.Name & "'"
    End If
End Sub

'@label:is-in-debug-mode
'@sub-title Test whether the workbook is in debug mode
'@return Boolean. True when debug mode is active.
Private Function InDebugMode() As Boolean
    InDebugMode = (StrComp(CStr(NamedRange(NAMEDEBUGMODE).Value), DEFAULTBOOLYES, vbTextCompare) = 0)
End Function

'@label:apply-protection-matrix
'@sub-title Apply stored protection settings to all listed worksheets
'@param wb Workbook. Target workbook whose sheets will be protected.
'@export
Private Sub ApplyProtectionMatrix(ByVal wb As Workbook)
    Dim Lo As ListObject
    Dim bodyRange As Range
    Dim rowIdx As Long
    Dim sheetName As String
    Dim allowShapesValue As Boolean
    Dim allowDeleteRows As Boolean


    Set Lo = ProtectedSheetsTable()
    Set bodyRange = Lo.DataBodyRange

    If bodyRange Is Nothing Then Exit Sub

    For rowIdx = 1 To bodyRange.Rows.Count
        sheetName = CStr(bodyRange.Cells(rowIdx, 1).Value)
        If LenB(sheetName) = 0 Then GoTo NextRow
        allowShapesValue = ParseBoolean(bodyRange.Cells(rowIdx, 2).Value)
        allowDeleteRows = ParseBoolean(bodyRange.Cells(rowIdx, 3).Value)

        On Error Resume Next
            ApplySheetProtection wb.Worksheets(sheetName), allowShapesValue, allowDeleteRows
        On Error GoTo 0
NextRow:
    Next rowIdx
End Sub

'@label:apply-sheet-protection
'@sub-title Protect a worksheet using debugging password and capability flags
Private Sub ApplySheetProtection(ByVal targetSheet As Worksheet, _
                                 ByVal allowShapes As Boolean, _
                                 ByVal allowDeleting As Boolean)


    Dim pwd As String
    pwd = TagValue(PasswordTagDebug)

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

'@label:protection-record
'@sub-title Retrieve stored protection preferences for a sheet
'@return BetterArray. Array(sheetName, allowShapes, allowDelete).
Private Function ProtectionRecord(ByVal sheetName As String) As BetterArray

    Dim Lo As ListObject
    Dim idRange As Range
    Dim found As Range
    Dim record As BetterArray

    Set Lo = ProtectedSheetsTable()
    Set idRange = Lo.ListColumns(1).DataBodyRange

    If idRange Is Nothing Then
        Set ProtectionRecord = EmptyProtectionRecord(sheetName)
        Exit Function
    End If

    On Error Resume Next
        Set found = idRange.Find(What:=sheetName, _
                                 LookAt:=xlWhole, _
                                 MatchCase:=True, _
                                 after:=idRange.Cells(idRange.Rows.Count, 1))
    On Error GoTo 0

    If found Is Nothing Then
        Set ProtectionRecord = EmptyProtectionRecord(sheetName)
        Exit Function
    End If

    Set record = New BetterArray
    record.LowerBound = 1
    record.Push sheetName
    record.Push CStr(found.Offset(0, 1).Value)
    record.Push CStr(found.Offset(0, 2).Value)

    Set ProtectionRecord = record
End Function

'@label:empty-protection-record
'@sub-title Construct default protection settings for a sheet
'@return BetterArray. Array(sheetName, "yes", "yes").
Private Function EmptyProtectionRecord(ByVal sheetName As String) As BetterArray

    Dim record As BetterArray
    Set record = New BetterArray
    record.LowerBound = 1
    record.Push sheetName
    record.Push DEFAULTBOOLYES
    record.Push DEFAULTBOOLYES
    Set EmptyProtectionRecord = record
End Function

'@label:persist-protection-settings
'@sub-title Insert or update the protection table entry for a sheet
Private Sub PersistProtectionSettings(ByVal sheetName As String, _
                                      ByVal allowShapes As Boolean, _
                                      ByVal allowDeletingRows As Boolean)


    Dim Lo As ListObject
    Dim idRange As Range
    Dim targetRow As Range

    Set Lo = ProtectedSheetsTable()

    Set idRange = Lo.ListColumns(1).DataBodyRange

    If idRange Is Nothing Then
        Set targetRow = Lo.ListRows.Add.Range
    Else
        On Error Resume Next
            Set targetRow = idRange.Find(What:=sheetName, LookAt:=xlWhole, MatchCase:=True)
        On Error GoTo 0
        If targetRow Is Nothing Then
            'Append a new row when the sheet was not found in the table.
            Set targetRow = Lo.ListRows.Add.Range
        End If
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
           "    Dim passObj As IPasswords" & vbNewLine & _
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
    Dim insertLine As Long
    Dim guardToken As String

    procedureName = "Workbook_BeforeClose"
    guardToken = "LeaveDebugModeOnClose"

    Set codeModule = component.CodeModule
    If codeModule Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, _
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
    Dim lineCount As Long

    Set codeModule = component.CodeModule

    If codeModule Is Nothing Then
        ThrowError ProjectError.ErrorUnexpectedState, "Unable to access code module for VB component"
    End If

    On Error Resume Next
        startLine = codeModule.ProcStartLine(procedureName, 0)
        If startLine > 0 Then
            endLine = codeModule.ProcCountLines(procedureName, 0)
            codeModule.DeleteLines startLine, endLine
        End If
    On Error GoTo 0

    lineCount = codeModule.CountOfLines
    If lineCount > 0 Then lineCount = lineCount + 1
    'Append the generated procedure at the end of the module.
    codeModule.InsertLines lineCount + 1, procedureCode
End Sub

'@label:copy-list-objects
'@sub-title Replicate ListObjects from source to target sheet
Private Sub CopyListObjects(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)

    Dim Lo As ListObject
    Dim targetRange As Range

    For Each Lo In sourceSheet.ListObjects
        Set targetRange = targetSheet.Range(Lo.Range.Address)
        'Copy values first so table resize includes data.
        targetRange.Value = Lo.Range.Value
        With targetSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
            .Name = Lo.Name
            .TableStyle = Lo.TableStyle
            .HeaderRowRange.Interior.Color = Lo.HeaderRowRange.Interior.Color
            .HeaderRowRange.Font.Color = Lo.HeaderRowRange.Font.Color
        End With
    Next lo
End Sub

'@label:copy-named-ranges
'@sub-title Copy critical named ranges to the target sheet
Private Sub CopyNamedRanges(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)

    Dim rangeNames As Variant
    Dim idx As Long
    Dim sourceRange As Range
    Dim targetRange As Range

    rangeNames = Array(NAMEDEBUGPASSWORD, NAMEPUBLICKEY, NAMELABPUBLICKEY, _
                       NAMEPRIVATEKEY, NAMELABPRIVATEKEY, NAMEDEBUGMODE, NAMEVERSION)

    For idx = LBound(rangeNames) To UBound(rangeNames)
        Set sourceRange = sourceSheet.Range(CStr(rangeNames(idx)))
        Set targetRange = targetSheet.Range(sourceRange.Address)
        targetRange.Value = sourceRange.Value
        If sourceRange.HasFormula Then targetRange.Formula = sourceRange.Formula
        On Error Resume Next
            targetSheet.Names(CStr(rangeNames(idx))).Delete
        On Error GoTo 0
        'Recreate the name so the clone is self-contained.
        targetSheet.Names.Add Name:=CStr(rangeNames(idx)), RefersTo:=targetRange, Visible:=False
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
