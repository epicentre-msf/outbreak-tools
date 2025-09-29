VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Passwords"

'@Interface
'@Folder("Security")
'@ModuleDescription("Interface exposing workbook password management behaviours")
'@depends Checking
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Public constants
'===============================================================================
'Byte-coded identifiers for Password tags to avoid Enum usage (macOS safe).
Public Enum PaswordTag
  PasswordTagDebug = 1
  PasswordTagPublic = 2
  PasswordTagPrivate = 3
  PasswordTagVersion = 4
  PasswordTagLabPublic = 5
  PasswordTagLabPrivate = 6
End Enum

'@section Factory
'===============================================================================
Public Function Create(ByVal passwordSheet As Worksheet) As IPasswords
End Function

'@section Core accessors
'===============================================================================
Public Property Get Sheet() As Worksheet
End Property

Public Property Get Workbook() As Workbook
End Property

'@details Retrieve the value stored in the named range associated with the supplied tag label.
'@param tagName Supported tags: debuggingpassword, publickey, privatekey, version, labpublickey, labprivatekey.
Public Property Get Value(ByVal tagName As String) As String
End Property

'@details Retrieve the value stored in the named range associated with the supplied tag identifier.
'@param tagId Byte identifier exposed as PasswordTag* constants.
Public Property Get TagValue(ByVal tagId As Byte) As String
End Property

'@details Exposes the range of a password-managed listobject.
'@param tableName Identifier of the table to retrieve (defaults to T_keys).
'@param includeHeaders Include header row when True.
Public Property Get TableRange(Optional ByVal tableName As String = "T_keys", _
                               Optional ByVal includeHeaders As Boolean = False) As Range
End Property

'@details Indicates whether any informational diagnostics have been captured.
Public Property Get HasCheckings() As Boolean
End Property

'@details Returns collected checking entries for diagnostic reporting.
Public Property Get CheckingValues() As IChecking
End Property

'@section Debug mode API
'===============================================================================
Public Sub EnterDebugMode(Optional ByVal targetWorkbook As Workbook)
End Sub

'@details Restore workbook protections previously removed by EnterDebugMode
Public Sub LeaveDebugMode(Optional ByVal targetWorkbook As Workbook)
End Sub

'@details Exit debug mode only when the flag indicates it is active.
Public Sub LeaveDebugModeIfActive(Optional ByVal targetWorkbook As Workbook)
End Sub

'@details Ensure the workbook exposes a named range referencing the protected sheets table.
Public Sub EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)
End Sub

'@details Inject before-close procedures that leave debug mode automatically.
Public Sub EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook)
End Sub

'@section Protection API
'===============================================================================
'@details Protect a worksheet according to supplied or stored rules.
'@param sheetObject Worksheet instance or name.
'@param allowShapes Allow manipulating drawing objects (optional).
'@param allowDeletingRows Allow deleting rows (optional).
'@param useTable When True, reuse stored table settings.
Public Sub ProtectSheet(ByVal sheetObject As Variant, _
                        Optional ByVal allowShapes As Variant, _
                        Optional ByVal allowDeletingRows As Variant, _
                        Optional ByVal useTable As Boolean = False)
End Sub

'@details Unprotect a worksheet using the debugging password.
'@param sheetName Worksheet name or "_active" for the active sheet.
Public Sub UnprotectSheet(Optional ByVal sheetName As String = "_active")
End Sub

'Compatibility wrappers (legacy API)
'@details Retained for backwards compatibility with legacy callers.
Public Sub Protect(Optional ByVal sheetObject As Variant, _
                   Optional ByVal allowShapes As Variant, _
                   Optional ByVal allowDeletingRows As Variant)
End Sub

'@details Retained for backwards compatibility with legacy callers.
Public Sub UnProtect(Optional ByVal sheetName As String = "_active")
End Sub

'@details Protect workbook structure using the debugging password.
Public Sub ProtectWorkbook(ByVal wb As Workbook)
End Sub

'@details Unprotect workbook structure using the debugging password.
Public Sub UnprotectWorkbook(ByVal wb As Workbook)
End Sub

Public Sub ProtectWkb(ByVal wb As Workbook, Optional ByVal protectSheetName As Boolean = True)
End Sub

Public Sub UnProtectWkb(ByVal wb As Workbook)
End Sub

'@section Keys
'===============================================================================
'@details Randomly select a public/private key pair and update the password worksheet.
Public Sub GenerateKey(ByVal trads As ITranslationObject)
End Sub

'@details Display the current private key to the user via a translated prompt.
Public Sub DisplayPrivateKey(ByVal trads As ITranslationObject)
End Sub

'@section Import/Export
'===============================================================================
'@details Export the password worksheet into another workbook, copying tables and named ranges.
Public Sub ExportToWorkbook(ByVal destination As Workbook)
End Sub

'@details Import keys and metadata from another password handler instance.
Public Sub ImportFrom(ByVal source As IPasswords)
End Sub

'@details Clone the password worksheet into another sheet and returns its handler.
Public Function CloneTo(ByVal targetSheet As Worksheet) As IPasswords
End Function

'Legacy wrappers
'@details Retained for backwards compatibility with legacy callers.
Public Sub Export(ByVal toWkb As Workbook)
End Sub

'@details Retained for backwards compatibility with legacy callers.
Public Sub Import(ByVal pass As IPasswords)
End Sub
