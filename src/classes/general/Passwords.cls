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
'@Folder("Security")
'@ModuleDescription("Lock/unlock workbooks and worksheets")
'@depends Checking, BetterArray
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IPasswords

'@section State
'===============================================================================
Private Type TPasswords
    sheet As Worksheet                       'Password worksheet backing the object
    workbook As Workbook                     'Cached workbook reference
    keyTable As ListObject                   'Cached list object holding keys
    protectedTable As ListObject             'Cached list object holding protection metadata
    rangeCache As BetterArray                'Cache for named ranges
    hasChecks As Boolean                     'Whether checkings were emitted
    checks As IChecking                      'Checking aggregation object
    checkCounter As Long                     'Incremental counter for checking keys
End Type

Private this As TPasswords

'@section Constants
'===============================================================================
Private Const CLASSNAME As String = "Passwords"
Private Const TABLE_KEYS As String = "T_keys"
Private Const TABLE_PROTECTED As String = "T_ProtectedSheets"
Private Const NAME_DEBUG_PASSWORD As String = "RNG_DebuggingPassword"
Private Const NAME_PUBLIC_KEY As String = "RNG_PublicKey"
Private Const NAME_LAB_PUBLIC_KEY As String = "RNG_LabPublicKey"
Private Const NAME_PRIVATE_KEY As String = "RNG_PrivateKey"
Private Const NAME_LAB_PRIVATE_KEY As String = "RNG_LabPrivateKey"
Private Const NAME_DEBUG_MODE As String = "RNG_DebugMode"
Private Const NAME_VERSION As String = "RNG_Version"
Private Const NAME_PROTECTED_SHEETS As String = "Passwords_ProtectedSheets"
Private Const DEFAULT_BOOL_YES As String = "yes"
Private Const DEFAULT_BOOL_NO As String = "no"
Private Const CHECKING_CONTEXT As String = "Passwords object trace"

'@section Tag identifiers (Byte-coded to avoid Enum usage on macOS)
'===============================================================================
Private Const TAG_DEBUG_PASSWORD As Byte = 1
Private Const TAG_PUBLIC_KEY As Byte = 2
Private Const TAG_PRIVATE_KEY As Byte = 3
Private Const TAG_VERSION As Byte = 4
Private Const TAG_LAB_PUBLIC_KEY As Byte = 5
Private Const TAG_LAB_PRIVATE_KEY As Byte = 6

'@section Factory
'===============================================================================
'@details Factory returning an interface reference that wraps the provided password worksheet.
'@param passwordSheet Worksheet hosting password tables and named ranges.
'@return Fully initialised password handler implementing IPasswords.
Public Function Create(ByVal passwordSheet As Worksheet) As IPasswords

    ValidateSheet passwordSheet

    With New Passwords
        Set .PasswordSheet = passwordSheet
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IPasswords
    '@details Expose the current instance through the IPasswords interface while ensuring caches are primed.
    '@return Current instance cast as IPasswords.
    EnsureCachesReady
    Set Self = Me
End Property

Friend Property Set PasswordSheet(ByVal passwordSheet As Worksheet)
'@details Internal setter used during construction to bind the password worksheet and initialise caches.
'@param passwordSheet Worksheet to associate with the handler.
    Set this.sheet = passwordSheet
    Set this.workbook = passwordSheet.Parent
    ResetCaches
    EnsureCachesReady
End Property

'@section Public properties
'===============================================================================
'@details Exposes the password sheet managed by this instance.
'@return Worksheet reference backing the password object.
Public Property Get Sheet() As Worksheet
    EnsureSheetReady
    Set Sheet = this.sheet
End Property

'@details Exposes the workbook owning the password worksheet.
'@return Workbook that contains the password sheet.
Public Property Get Workbook() As Workbook
    EnsureWorkbookReady
    Set Workbook = this.workbook
End Property

'@details Retrieve the value stored in the named range associated with the supplied tag.
'@param tagName Friendly tag identifier (debuggingpassword, publickey, privatekey, version, labpublickey, labprivatekey).
'@return Value stored in the named range associated with the tag.
Public Property Get Value(ByVal tagName As String) As String
    Value = TagValue(TagFromName(tagName))
End Property

'@details Retrieve the value stored in a named range identified by the byte-coded tag.
'@param tagId Byte-coded identifier defined on IPasswords (PasswordTag* constants).
'@return Value stored in the named range corresponding to the tag identifier.
Public Property Get TagValue(ByVal tagId As Byte) As String
    Dim targetRange As Range
    Set targetRange = NamedRange(TagToRangeName(tagId))
    TagValue = CStr(targetRange.Value)
End Property

'@details Retrieve the range of a password-managed table.
'@param tableName Table identifier (T_keys or T_ProtectedSheets).
'@param includeHeaders When True, include header row in the returned range.
'@return Range representing the requested listobject on the password sheet.
Public Property Get TableRange(Optional ByVal tableName As String = TABLE_KEYS, _
                               Optional ByVal includeHeaders As Boolean = False) As Range

    Dim lo As ListObject

    Select Case LCase$(Trim$(tableName))
        Case LCase$(TABLE_KEYS)
            Set lo = KeysTable()
        Case LCase$(TABLE_PROTECTED)
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

Public Property Get HasCheckings() As Boolean
    '@details Indicates whether any informational diagnostics were recorded.
    '@return True when at least one checking entry exists.
    HasCheckings = this.hasChecks
End Property

Public Property Get CheckingValues() As IChecking
    '@details Retrieve the aggregated checking entries for downstream reporting.
    '@return Checking object aggregating informational messages (Nothing when no entries exist).
    If this.hasChecks Then
        Set CheckingValues = this.checks
    End If
End Property

'@section Debug mode management
'===============================================================================
'@details Puts the target workbook into a fully unprotected state, updating the debug mode flag and logging the transition.
'@param targetWorkbook Optional workbook to toggle into debug mode (defaults to password workbook).
Public Sub EnterDebugMode(Optional ByVal targetWorkbook As Workbook)

    Dim wb As Workbook
    Dim childSheet As Worksheet

    Set wb = ResolveWorkbook(targetWorkbook)

    'Unlock workbook structure before processing child worksheets.
    UnprotectWorkbook wb

    For Each childSheet In wb.Worksheets
        'Each sheet must be explicitly unprotected for unrestricted debugging.
        UnprotectSheet childSheet.Name
    Next childSheet

    NamedRange(NAME_DEBUG_MODE).Value = DEFAULT_BOOL_YES
    EnsureProtectedSheetsName wb
    LogInfo "Workbook entered debug mode; protections removed", checkingWarning
End Sub

'@details Restores workbook structure protection and worksheet protections per stored policy.
'@param targetWorkbook Optional workbook to restore protections for (defaults to password workbook).
Public Sub LeaveDebugMode(Optional ByVal targetWorkbook As Workbook)

    Dim wb As Workbook

    Set wb = ResolveWorkbook(targetWorkbook)

    'Reapply worksheet-level protections captured in the metadata table.
    ApplyProtectionMatrix wb
    'Restore workbook structure protection once sheets are locked again.
    ProtectWorkbook wb
    NamedRange(NAME_DEBUG_MODE).Value = DEFAULT_BOOL_NO
    LogInfo "Workbook protections restored after leaving debug mode"
End Sub

'@details Convenience wrapper that silently exits when debug mode is already off.
'@param targetWorkbook Optional workbook whose debug mode status should be evaluated.
Public Sub LeaveDebugModeIfActive(Optional ByVal targetWorkbook As Workbook)
    If InDebugMode() Then
        LeaveDebugMode ResolveWorkbook(targetWorkbook)
    End If
End Sub

'@details Ensures the workbook exposes a named range referencing the protected sheets table.
'@param targetWorkbook Optional workbook receiving or updating the named reference (defaults to password workbook).
Public Sub EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)

    Dim wb As Workbook
    Dim namedItem As Name
    Dim targetRange As Range
    Dim refersTo As String

    Set wb = ResolveWorkbook(targetWorkbook)
    Set targetRange = ProtectedSheetsTable().Range

    refersTo = "=" & targetRange.Address(True, True, xlA1, True)

    On Error Resume Next
        Set namedItem = wb.Names(NAME_PROTECTED_SHEETS)
    On Error GoTo 0

    If namedItem Is Nothing Then
        'Create the name so closing handlers can look it up reliably.
        wb.Names.Add Name:=NAME_PROTECTED_SHEETS, RefersTo:=refersTo
    ElseIf StrComp(namedItem.RefersTo, refersTo, vbBinaryCompare) <> 0 Then
        'Refresh the reference if the table size changed.
        namedItem.RefersTo = refersTo
    End If
End Sub

'@details Injects before-close safeguards that automatically exit debug mode during workbook shutdown.
'@param targetWorkbook Optional workbook where the guard procedures should be written (defaults to password workbook).
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
    InsertOrReplaceProcedure vbComp, "Workbook_BeforeClose", BuildBeforeCloseProcedure
    Exit Sub

NoProject:
    ThrowError ProjectError.ObjectNotInitialized, _
              "Unable to access VBProject for workbook '" & wb.Name & "'"
    Exit Sub

WorkbookModuleMissing:
    ThrowError ProjectError.ElementNotFound, _
              "Workbook module '" & wb.CodeName & "' is missing"
End Sub


'@section Worksheet protection management
'===============================================================================
'@details Protect a worksheet with the debugging password and either supplied or stored options.
'@param sheetObject Worksheet instance or sheet name to protect.
'@param allowShapes Optional Boolean permitting shape manipulation (defaults to False when not supplied).
'@param allowDeletingRows Optional Boolean enabling row deletion (defaults to True when not supplied).
'@param useTable When True, reuse stored table settings instead of supplied arguments.
Public Sub ProtectSheet(ByVal sheetObject As Variant, _
                        Optional ByVal allowShapes As Variant, _
                        Optional ByVal allowDeletingRows As Variant, _
                        Optional ByVal useTable As Boolean = False)

    Dim sheetName As String
    Dim targetSheet As Worksheet
    Dim allowDraw As Boolean
    Dim allowDelete As Boolean
    Dim record As BetterArray

    EnsureWorkbookReady
    sheetName = ResolveSheetName(sheetObject)

    If InDebugMode() Then
        LogInfo "Protection request ignored because workbook is in debug mode", checkingWarning
        Exit Sub
    End If

    On Error Resume Next
        Set targetSheet = this.workbook.Worksheets(sheetName)
    On Error GoTo 0

    If targetSheet Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Worksheet '" & sheetName & "' not found"
    End If

    If useTable Then
        'Reuse existing stored preferences to guarantee consistent application.
        Set record = ProtectionRecord(sheetName)
        allowDraw = ParseBoolean(record.Item(record.LowerBound + 1))
        allowDelete = ParseBoolean(record.Item(record.LowerBound + 2))
    Else
        'Interpret optional arguments and persist the choice for future reuse.
        allowDraw = CoerceBooleanOptional(allowShapes, False)
        allowDelete = CoerceBooleanOptional(allowDeletingRows, True)
        PersistProtectionSettings sheetName, allowDraw, allowDelete
    End If

    ApplySheetProtection targetSheet, allowDraw, allowDelete
End Sub

'@details Unlocks a worksheet using the stored debugging password.
'@param sheetName Worksheet name to unlock; "_active" targets the current active sheet.
Public Sub UnprotectSheet(Optional ByVal sheetName As String = "_active")

    Dim targetSheet As Worksheet
    Dim pwd As String

    EnsureWorkbookReady

    If StrComp(sheetName, "_active", vbBinaryCompare) = 0 Then
        Set targetSheet = ActiveSheet
    Else
        On Error Resume Next
            Set targetSheet = this.workbook.Worksheets(sheetName)
        On Error GoTo 0
    End If

    If targetSheet Is Nothing Then Exit Sub

    pwd = TagValue(TAG_DEBUG_PASSWORD)

    On Error Resume Next
        If targetSheet.ProtectContents Then targetSheet.Unprotect pwd
    On Error GoTo 0
End Sub

'@details Protects workbook structure with the debugging password to prevent sheet renames.
'@param wb Workbook whose structure should be protected.
Public Sub ProtectWorkbook(ByVal wb As Workbook)
    Dim pwd As String
    pwd = TagValue(TAG_DEBUG_PASSWORD)
    wb.Protect Password:=pwd, Structure:=True
End Sub

'@details Unlocks workbook structure using the debugging password.
'@param wb Workbook whose structure should be unlocked.
Public Sub UnprotectWorkbook(ByVal wb As Workbook)
    Dim pwd As String
    pwd = TagValue(TAG_DEBUG_PASSWORD)
    On Error Resume Next
        wb.Unprotect Password:=pwd
    On Error GoTo 0
End Sub

'@section Key management
'===============================================================================
'@details Randomly selects a public/private key pair from the keys table and updates the worksheet.
'@param trads Translation helper used to localise informational prompts.
Public Sub GenerateKey(ByVal trads As ITranslationObject)

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

    Randomize Timer
    recordIndex = CLng((dataRange.Rows.Count * Rnd) + 1)
    If recordIndex > dataRange.Rows.Count Then recordIndex = dataRange.Rows.Count
    If recordIndex < 1 Then recordIndex = 1

    publicValue = CStr(dataRange.Cells(recordIndex, 1).Value)
    privateValue = CStr(dataRange.Cells(recordIndex, 2).Value)

    NamedRange(NAME_PUBLIC_KEY).Value = publicValue
    NamedRange(NAME_PRIVATE_KEY).Value = privateValue

    DisplayPrivateKey trads
End Sub

'@details Displays the current private key to the user through a translated prompt.
'@param trads Translation helper used to localise the message box.
Public Sub DisplayPrivateKey(ByVal trads As ITranslationObject)
    If trads Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Translation object is required to display private key"
    End If

    MsgBox trads.TranslatedValue("MSG_Password") & TagValue(TAG_PRIVATE_KEY), vbInformation, trads.TranslatedValue("MSG_Title")
End Sub

'@section Import/Export
'===============================================================================
'@details Exports the password sheet, copying tables and named ranges into another workbook.
'@param destination Target workbook receiving a copy of the password sheet.
Public Sub ExportToWorkbook(ByVal destination As Workbook)

    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet

    EnsureWorkbookReady

    If destination Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Destination workbook is required for export"
    End If

    Set sourceSheet = Sheet

    On Error Resume Next
        Set targetSheet = destination.Worksheets(sourceSheet.Name)
    On Error GoTo 0

    If Not targetSheet Is Nothing Then
        ThrowError ProjectError.ElementShouldNotExists, _
                  "Workbook already contains a worksheet named '" & sourceSheet.Name & "'"
    End If

    Set targetSheet = destination.Worksheets.Add(After:=destination.Worksheets(destination.Worksheets.Count))
    targetSheet.Name = sourceSheet.Name

    CopyListObjects sourceSheet, targetSheet
    CopyNamedRanges sourceSheet, targetSheet
    targetSheet.Visible = xlSheetVeryHidden
End Sub

'@details Imports keys and metadata from another password object instance.
'@param source Password object providing data for import.
Public Sub ImportFrom(ByVal source As IPasswords)

    Dim loSource As Range
    Dim destinationLo As ListObject
    Dim resizedRange As Range

    If source Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source password object is required for import"
    End If

    NamedRange(NAME_PUBLIC_KEY).Value = source.Value("publickey")
    NamedRange(NAME_PRIVATE_KEY).Value = source.Value("privatekey")
    NamedRange(NAME_VERSION).Value = source.Value("version")

    Set loSource = source.TableRange(TABLE_KEYS, includeHeaders:=False)
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

'@details Creates a structural copy of the password worksheet on another sheet and returns its handler.
'@param targetSheet Worksheet that will receive a cloned password setup.
'@return Newly instantiated password handler bound to the cloned sheet.
Public Function CloneTo(ByVal targetSheet As Worksheet) As IPasswords

    If targetSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target sheet is required when cloning password data"
    End If

    targetSheet.Cells.Clear
    CopyListObjects Sheet, targetSheet
    CopyNamedRanges Sheet, targetSheet

    With Create(targetSheet)
        .EnsureProtectedSheetsName targetSheet.Parent
        Set CloneTo = .Self
    End With
End Function

'@section Interface implementation
'===============================================================================
Private Property Get IPasswords_Sheet() As Worksheet
    Set IPasswords_Sheet = Sheet
End Property

Private Property Get IPasswords_Workbook() As Workbook
    Set IPasswords_Workbook = Workbook
End Property

Private Property Get IPasswords_Value(ByVal tagName As String) As String
    IPasswords_Value = Value(tagName)
End Property

Private Property Get IPasswords_TagValue(ByVal tagId As Byte) As String
    IPasswords_TagValue = TagValue(tagId)
End Property

Private Property Get IPasswords_TableRange(Optional ByVal tableName As String = TABLE_KEYS, _
                                           Optional ByVal includeHeaders As Boolean = False) As Range
    Set IPasswords_TableRange = TableRange(tableName, includeHeaders)
End Property

Private Property Get IPasswords_HasCheckings() As Boolean
    IPasswords_HasCheckings = HasCheckings
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

Private Sub IPasswords_LeaveDebugModeIfActive(Optional ByVal targetWorkbook As Workbook)
    LeaveDebugModeIfActive targetWorkbook
End Sub

Private Sub IPasswords_EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)
    EnsureProtectedSheetsName targetWorkbook
End Sub

Private Sub IPasswords_EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook)
    EnsureDebugExitHandler targetWorkbook
End Sub

Private Sub IPasswords_ProtectSheet(ByVal sheetObject As Variant, _
                                     Optional ByVal allowShapes As Variant, _
                                     Optional ByVal allowDeletingRows As Variant, _
                                     Optional ByVal useTable As Boolean = False)
    ProtectSheet sheetObject, allowShapes, allowDeletingRows, useTable
End Sub

Private Sub IPasswords_UnprotectSheet(Optional ByVal sheetName As String = "_active")
    UnprotectSheet sheetName
End Sub

Private Sub IPasswords_Protect(Optional ByVal sheetObject As Variant, _
                               Optional ByVal allowShapes As Variant, _
                               Optional ByVal allowDeletingRows As Variant)
    'Legacy entry point maintained for backward compatibility.
    ProtectSheet sheetObject, allowShapes, allowDeletingRows
End Sub

Private Sub IPasswords_UnProtect(Optional ByVal sheetName As String)
    'Legacy entry point maintained for backward compatibility.
    UnprotectSheet sheetName
End Sub

Private Sub IPasswords_ProtectWorkbook(ByVal wb As Workbook)
    ProtectWorkbook wb
End Sub

Private Sub IPasswords_UnprotectWorkbook(ByVal wb As Workbook)
    UnprotectWorkbook wb
End Sub

Private Sub IPasswords_ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
    Dim pwd As String
    pwd = TagValue(TAG_DEBUG_PASSWORD)
    wb.Protect Password:=pwd, Structure:=protectSheetName
End Sub

Private Sub IPasswords_UnProtectWkb(ByVal wb As Workbook)
    Dim pwd As String
    pwd = TagValue(TAG_DEBUG_PASSWORD)
    wb.Unprotect Password:=pwd
End Sub

Private Sub IPasswords_GenerateKey(ByVal trads As ITranslationObject)
    GenerateKey trads
End Sub

Private Sub IPasswords_DisplayPrivateKey(ByVal trads As ITranslationObject)
    DisplayPrivateKey trads
End Sub

Private Sub IPasswords_ExportToWorkbook(ByVal destination As Workbook)
    ExportToWorkbook destination
End Sub

Private Sub IPasswords_ImportFrom(ByVal source As IPasswords)
    ImportFrom source
End Sub

Private Function IPasswords_CloneTo(ByVal targetSheet As Worksheet) As IPasswords
    Set IPasswords_CloneTo = CloneTo(targetSheet)
End Function

Private Sub IPasswords_Export(ByVal toWkb As Workbook)
    'Legacy entry point maintained for backward compatibility.
    ExportToWorkbook toWkb
End Sub

Private Sub IPasswords_Import(ByVal pass As IPasswords)
    'Legacy entry point maintained for backward compatibility.
    ImportFrom pass
End Sub

'@section Private helpers - Validation & ensuring state
'===============================================================================
Private Sub ValidateSheet(ByVal passwordSheet As Worksheet)

    '@details Validates that the provided worksheet contains the structures required for password management.

    If passwordSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Password sheet is required"
    End If

    If passwordSheet.ListObjects(TABLE_KEYS) Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Keys table '" & TABLE_KEYS & "' is missing"
    End If

    If passwordSheet.ListObjects(TABLE_PROTECTED) Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Protected sheets table '" & TABLE_PROTECTED & "' is missing"
    End If

    EnsureRangeExists passwordSheet, NAME_DEBUG_PASSWORD
    EnsureRangeExists passwordSheet, NAME_PUBLIC_KEY
    EnsureRangeExists passwordSheet, NAME_PRIVATE_KEY
    EnsureRangeExists passwordSheet, NAME_DEBUG_MODE
    EnsureRangeExists passwordSheet, NAME_VERSION
End Sub

Private Sub EnsureSheetReady()
    '@details Guards against use of the handler before a worksheet has been associated with it.
    If this.sheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Password sheet is not initialised"
    End If
End Sub

Private Sub EnsureWorkbookReady()
    '@details Ensures the workbook reference is resolved prior to operations that require it.
    EnsureSheetReady
    If this.workbook Is Nothing Then
        Set this.workbook = this.sheet.Parent
    End If
End Sub

Private Sub EnsureCachesReady()
    '@details Lazily instantiate reusable caches to avoid repeated worksheet scans.
    If this.rangeCache Is Nothing Then
        Set this.rangeCache = New BetterArray
        this.rangeCache.LowerBound = 1
    End If
End Sub

Private Sub ResetCaches()
    '@details Clears cached listobjects and named ranges after structural updates.
    Set this.rangeCache = Nothing
    Set this.keyTable = Nothing
    Set this.protectedTable = Nothing
End Sub

Private Function KeysTable() As ListObject
    '@details Retrieves and caches the keys listobject, ensuring it exists on the password sheet.
    EnsureSheetReady
    If this.keyTable Is Nothing Then
        Set this.keyTable = this.sheet.ListObjects(TABLE_KEYS)
        If this.keyTable Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                      "ListObject '" & TABLE_KEYS & "' not found on sheet '" & this.sheet.Name & "'"
        End If
    End If
    Set KeysTable = this.keyTable
End Function

Private Function ProtectedSheetsTable() As ListObject
    '@details Retrieves and caches the protected-sheets table.
    EnsureSheetReady
    If this.protectedTable Is Nothing Then
        Set this.protectedTable = this.sheet.ListObjects(TABLE_PROTECTED)
        If this.protectedTable Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                      "ListObject '" & TABLE_PROTECTED & "' not found on sheet '" & this.sheet.Name & "'"
        End If
    End If
    Set ProtectedSheetsTable = this.protectedTable
End Function

Private Function NamedRange(ByVal rangeName As String) As Range

    Dim cached As Variant
    Dim idx As Long
    Dim entry As Variant
    Dim result As Range

    EnsureSheetReady
    EnsureCachesReady

    For idx = this.rangeCache.LowerBound To this.rangeCache.UpperBound
        entry = this.rangeCache.Item(idx)
        If IsArray(entry) Then
            If StrComp(CStr(entry(0)), LCase$(rangeName), vbBinaryCompare) = 0 Then
                Set NamedRange = entry(1)
                Exit Function
            End If
        End If
    Next idx

    On Error Resume Next
        Set result = this.sheet.Range(rangeName)
    On Error GoTo 0

    If result Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & rangeName & "' is missing on sheet '" & this.sheet.Name & "'"
    End If

    this.rangeCache.Push Array(LCase$(rangeName), result)
    Set NamedRange = result
End Function

Private Function ResolveWorkbook(Optional ByVal candidate As Workbook) As Workbook
    '@details Returns the supplied workbook or falls back to the workbook backing the password sheet.
    If candidate Is Nothing Then
        EnsureWorkbookReady
        Set ResolveWorkbook = this.workbook
    Else
        Set ResolveWorkbook = candidate
    End If
End Function

Private Function ResolveSheetName(ByVal sheetObject As Variant) As String

    Dim typeNameValue As String

    typeNameValue = TypeName(sheetObject)

    Select Case typeNameValue
        Case "Worksheet"
            ResolveSheetName = CStr(sheetObject.Name)
        Case "String"
            ResolveSheetName = CStr(sheetObject)
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "Expected worksheet or sheet name, received " & typeNameValue
    End Select
End Function

Private Function TagToRangeName(ByVal tagId As Byte) As String
    '@details Maps a tag identifier to the corresponding named range identifier.
    Select Case tagId
        Case TAG_DEBUG_PASSWORD: TagToRangeName = NAME_DEBUG_PASSWORD
        Case TAG_PUBLIC_KEY: TagToRangeName = NAME_PUBLIC_KEY
        Case TAG_PRIVATE_KEY: TagToRangeName = NAME_PRIVATE_KEY
        Case TAG_VERSION: TagToRangeName = NAME_VERSION
        Case TAG_LAB_PUBLIC_KEY: TagToRangeName = NAME_LAB_PUBLIC_KEY
        Case TAG_LAB_PRIVATE_KEY: TagToRangeName = NAME_LAB_PRIVATE_KEY
        Case Else
            ThrowError ProjectError.InvalidArgument, "Unknown password tag identifier: " & CStr(tagId)
    End Select
End Function

Private Function TagFromName(ByVal tagName As String) As Byte
    '@details Converts a friendly tag label into the byte-coded identifier used internally.

    Select Case LCase$(Trim$(tagName))
        Case "debuggingpassword": TagFromName = TAG_DEBUG_PASSWORD
        Case "publickey": TagFromName = TAG_PUBLIC_KEY
        Case "privatekey": TagFromName = TAG_PRIVATE_KEY
        Case "version": TagFromName = TAG_VERSION
        Case "labpublickey": TagFromName = TAG_LAB_PUBLIC_KEY
        Case "labprivatekey": TagFromName = TAG_LAB_PRIVATE_KEY
        Case Else
            ThrowError ProjectError.InvalidArgument, "Unknown password tag '" & tagName & "'"
    End Select
End Function

Private Sub EnsureRangeExists(ByVal targetSheet As Worksheet, ByVal rangeName As String)
    '@details Validates that a required named range exists on the password worksheet.
    Dim rng As Range
    On Error Resume Next
        Set rng = targetSheet.Range(rangeName)
    On Error GoTo 0
    If rng Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Named range '" & rangeName & "' is missing on sheet '" & targetSheet.Name & "'"
    End If
End Sub

Private Function InDebugMode() As Boolean
    '@details Tests whether the password sheet declares the workbook in debug mode.
    InDebugMode = (StrComp(CStr(NamedRange(NAME_DEBUG_MODE).Value), DEFAULT_BOOL_YES, vbTextCompare) = 0)
End Function

Private Sub ApplyProtectionMatrix(ByVal wb As Workbook)

    '@details Applies stored protection settings to each worksheet listed in the protected-sheets table.

    Dim lo As ListObject
    Dim bodyRange As Range
    Dim rowIdx As Long
    Dim sheetName As String
    Dim allowShapesValue As Boolean
    Dim allowDeleteValue As Boolean
    Dim record As BetterArray

    Set lo = ProtectedSheetsTable()
    Set bodyRange = lo.DataBodyRange

    If bodyRange Is Nothing Then Exit Sub

    For rowIdx = 1 To bodyRange.Rows.Count
        sheetName = CStr(bodyRange.Cells(rowIdx, 1).Value)
        If LenB(sheetName) = 0 Then GoTo NextRow

        Set record = New BetterArray
        record.LowerBound = 1
        record.Push sheetName
        record.Push CStr(bodyRange.Cells(rowIdx, 2).Value)
        record.Push CStr(bodyRange.Cells(rowIdx, 3).Value)

        allowShapesValue = ParseBoolean(record.Item(record.LowerBound + 1))
        allowDeleteValue = ParseBoolean(record.Item(record.LowerBound + 2))

        On Error Resume Next
            ApplySheetProtection wb.Worksheets(sheetName), allowShapesValue, allowDeleteValue
        On Error GoTo 0
NextRow:
    Next rowIdx
End Sub

Private Sub ApplySheetProtection(ByVal targetSheet As Worksheet, _
                                 ByVal allowShapes As Boolean, _
                                 ByVal allowDeleting As Boolean)

    '@details Protects a worksheet using the debugging password and supplied capability flags.

    Dim pwd As String
    pwd = TagValue(TAG_DEBUG_PASSWORD)

    targetSheet.Protect Password:=pwd, _
        DrawingObjects:=Not allowShapes, _
        Contents:=True, _
        Scenarios:=True, _
        AllowInsertingRows:=True, _
        AllowSorting:=False, _
        AllowFiltering:=True, _
        AllowFormattingCells:=False, _
        AllowDeletingRows:=allowDeleting, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True
End Sub

Private Function ProtectionRecord(ByVal sheetName As String) As BetterArray

    '@details Retrieve stored protection preferences for the specified sheet, returning defaults when absent.

    Dim lo As ListObject
    Dim idRange As Range
    Dim found As Range
    Dim record As BetterArray

    Set lo = ProtectedSheetsTable()
    Set idRange = lo.ListColumns(1).DataBodyRange

    If idRange Is Nothing Then
        Set ProtectionRecord = EmptyProtectionRecord(sheetName)
        Exit Function
    End If

    On Error Resume Next
        Set found = idRange.Find(What:=sheetName, LookAt:=xlWhole, MatchCase:=True)
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

Private Function EmptyProtectionRecord(ByVal sheetName As String) As BetterArray
    '@details Constructs a BetterArray initialised with default protection settings.
    Dim record As BetterArray
    Set record = New BetterArray
    record.LowerBound = 1
    record.Push sheetName
    record.Push DEFAULT_BOOL_NO
    record.Push DEFAULT_BOOL_YES
    Set EmptyProtectionRecord = record
End Function

Private Sub PersistProtectionSettings(ByVal sheetName As String, _
                                      ByVal allowShapes As Boolean, _
                                      ByVal allowDeletingRows As Boolean)

    '@details Inserts or updates the protection table entry for the specified sheet.

    Dim lo As ListObject
    Dim idRange As Range
    Dim targetRow As Range

    Set lo = ProtectedSheetsTable()

    Set idRange = lo.ListColumns(1).DataBodyRange

    If idRange Is Nothing Then
        Set targetRow = lo.ListRows.Add.Range
    Else
        On Error Resume Next
            Set targetRow = idRange.Find(What:=sheetName, LookAt:=xlWhole, MatchCase:=True)
        On Error GoTo 0
        If targetRow Is Nothing Then
            'Append a new row when the sheet was not found in the table.
            Set targetRow = lo.ListRows.Add.Range
        End If
    End If

    targetRow.Cells(1, 1).Value = sheetName
    targetRow.Cells(1, 2).Value = BooleanToFlag(allowShapes)
    targetRow.Cells(1, 3).Value = BooleanToFlag(allowDeletingRows)
End Sub

Private Function BooleanToFlag(ByVal value As Boolean) As String
    '@details Converts a Boolean to the "yes"/"no" textual flag used in the protection table.
    If value Then
        BooleanToFlag = DEFAULT_BOOL_YES
    Else
        BooleanToFlag = DEFAULT_BOOL_NO
    End If
End Function

Private Function ParseBoolean(ByVal value As Variant) As Boolean
    '@details Interprets stored textual flags where any value other than "no" is treated as True.
    ParseBoolean = (StrComp(CStr(value), DEFAULT_BOOL_NO, vbTextCompare) <> 0)
End Function

Private Function CoerceBooleanOptional(ByVal candidate As Variant, ByVal defaultValue As Boolean) As Boolean
    '@details Evaluates optional arguments that may be missing, null, or explicitly provided.
    If IsMissing(candidate) Then
        CoerceBooleanOptional = defaultValue
    ElseIf IsNull(candidate) Then
        CoerceBooleanOptional = defaultValue
    Else
        CoerceBooleanOptional = CBool(candidate)
    End If
End Function

Private Function BuildLeaveDebugProcedure() As String
    '@details Generates VBA source code that exits debug mode safely when the workbook closes.
    Dim code As String
    code = "Private Sub LeaveDebugModeOnClose()" & vbNewLine & _
           "    On Error Resume Next" & vbNewLine & _
           "    Dim guardName As Name" & vbNewLine & _
           "    Dim targetSheet As Worksheet" & vbNewLine & _
           "    Set guardName = ThisWorkbook.Names(""" & NAME_PROTECTED_SHEETS & """)" & vbNewLine & _
           "    If guardName Is Nothing Then Exit Sub" & vbNewLine & _
           "    Set targetSheet = guardName.RefersToRange.Worksheet" & vbNewLine & _
           "    Dim passObj As IPasswords" & vbNewLine & _
           "    Set passObj = Passwords.Create(targetSheet)" & vbNewLine & _
           "    passObj.LeaveDebugModeIfActive ThisWorkbook" & vbNewLine & _
           "End Sub"
    BuildLeaveDebugProcedure = code
End Function

Private Function BuildBeforeCloseProcedure() As String
    '@details Generates VBA source code wiring the workbook BeforeClose event to the debug exit helper.
    Dim code As String
    code = "Private Sub Workbook_BeforeClose(ByVal Cancel As Boolean)" & vbNewLine & _
           "    On Error Resume Next" & vbNewLine & _
           "    LeaveDebugModeOnClose" & vbNewLine & _
           "End Sub"
    BuildBeforeCloseProcedure = code
End Function

Private Sub InsertOrReplaceProcedure(ByVal component As Object, _
                                     ByVal procedureName As String, _
                                     ByVal procedureCode As String)

    '@details Inserts the supplied code into the component, replacing any existing procedure with the same name.

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

'@details Replicates every listobject from the source sheet into the target sheet, preserving style and name.
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
        End With
    Next lo
End Sub

'@details Copies critical named ranges to the target sheet, preserving formulas and refreshing definitions.
Private Sub CopyNamedRanges(ByVal sourceSheet As Worksheet, ByVal targetSheet As Worksheet)

    Dim rangeNames As Variant
    Dim idx As Long
    Dim sourceRange As Range
    Dim targetRange As Range

    rangeNames = Array(NAME_DEBUG_PASSWORD, NAME_PUBLIC_KEY, NAME_LAB_PUBLIC_KEY, _
                       NAME_PRIVATE_KEY, NAME_LAB_PRIVATE_KEY, NAME_DEBUG_MODE, NAME_VERSION)

    For idx = LBound(rangeNames) To UBound(rangeNames)
        Set sourceRange = sourceSheet.Range(CStr(rangeNames(idx)))
        Set targetRange = targetSheet.Range(sourceRange.Address)
        targetRange.Value = sourceRange.Value
        If LenB(CStr(sourceRange.Formula)) > 0 Then
            targetRange.Formula = sourceRange.Formula
        End If
        On Error Resume Next
            targetSheet.Names(CStr(rangeNames(idx))).Delete
        On Error GoTo 0
        'Recreate the name so the clone is self-contained.
        targetSheet.Names.Add Name:=CStr(rangeNames(idx)), RefersTo:=targetRange
    Next idx
End Sub

 '@details Append an informational entry to the internal checking log, initialising it on first use.
'@param label Message describing the event.
'@param scope Checking severity classification (note/warning/error).
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    If this.checks Is Nothing Then
        Set this.checks = Checking.Create(CHECKING_CONTEXT)
    End If

    this.checks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    '@details Raises a ProjectError-coded runtime error scoped to the Passwords class.
    Err.Raise errNumber, CLASSNAME, message
End Sub