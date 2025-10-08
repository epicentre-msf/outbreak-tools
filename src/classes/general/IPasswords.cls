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
'@Folder("General")
'@ModuleDescription("Interface of Passwords")
'@depends Checking
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit


'@section Tag identifiers
'===============================================================================
Public Enum PaswordTag
  PasswordTagDebug = 1
  PasswordTagPublic = 2
  PasswordTagPrivate = 3
  PasswordTagVersion = 4
  PasswordTagMode = 5
  PasswordTagLabPublic = 6
  PasswordTagLabPrivate = 7
End Enum

'@jump:Create
'@section Core accessors
'===============================================================================

'@jump:passwordsheet
'@sub-title Worksheet of the password object
Public Property Get PasswordSheet() As Worksheet: End Property

'@jump:value
'@sub-title Retrieve the value stored in the named range associated with the supplied tag label.
'@param tagName Supported tags: debuggingpassword, publickey, privatekey, version, labpublickey, labprivatekey.
Public Property Get Value(ByVal tagName As String) As String: End Property

'@sub-title Control whether UI prompts are displayed during operations.
Public Property Get DisplayPrompts() As Boolean: End Property

Public Property Let DisplayPrompts(ByVal value As Boolean) : End Property

'@sub-title Retrieve the most recent private key prompt text and title.
Public Property Get LastPrivatePrompt() As String: End Property
Public Property Get LastPrivatePromptTitle() As String: End Property

'@jump:tablerange
'@sub-title Public/Private keys table
Public Property Get TableRange(ByVal tableName As String, Optional ByVal includeHeaders As Boolean = False) As Range: End Property

'@jump:hascheckings
'@sub-title Indicates whether any informational diagnostics have been captured.
Public Property Get HasCheckings() As Boolean: End Property

'@jump:checkingvalues
'@sub-title Returns collected checking entries for diagnostic reporting.
Public Property Get CheckingValues() As IChecking: End Property

Public Property Get IsInDebugMode() As Boolean : End Property


'@section Debug mode API
'===============================================================================
Public Sub EnterDebugMode(Optional ByVal targetWorkbook As Workbook): End Sub

'@sub-title Restore workbook protections previously removed by EnterDebugMode
Public Sub LeaveDebugMode(Optional ByVal targetWorkbook As Workbook): End Sub

'@sub-title Inject before-close procedures that leave debug mode automatically.
Public Sub EnsureDebugExitHandler(Optional ByVal targetWorkbook As Workbook): 
End Sub

'@jump:ProtectionApi
'@section Protection API
'===============================================================================

'@sub-title Protect a worksheet according to supplied or stored rules.
'@param targetObject A Variant. Could be a Workbook, a Worksheet or a String, the name of the worksheet to protect. "_active"
'will protect the active worksheet. It is the default behavior. "_wbactive" will protect the active workbook when supplied.
'@param allowShapes Allow manipulating drawing objects (optional).
'@param allowDeletingRows Allow deleting rows (optional).
'@param registerState When True, reuse stored table settings or create worksheet state.
Public Sub Protect(Optional targetObject As Variant = "_active", _
                   Optional ByVal allowShapes As Boolean = True, _
                   Optional ByVal allowDeletingRows As Boolean = True, _
                   Optional ByVal registerState As Boolean = True)
End Sub

Public Sub UnProtect(Optional targetObject As Variant = "_active")
End Sub

Public Sub EnsureProtectedSheetsName(Optional ByVal targetWorkbook As Workbook)
End Sub


'@section Keys
'===============================================================================
'@sub-title Randomly select a public/private key pair and update the password worksheet.
Public Sub GenerateKey(ByVal trads As ITranslationObject): End Sub

'@sub-title Display the current private key to the user via a translated prompt.
Public Sub DisplayPrivateKey(ByVal trads As ITranslationObject): End Sub

'@section Import/Export
'===============================================================================
'@sub-title Export the password worksheet into another workbook, copying tables and named ranges.
Public Sub ExportToWorkbook(ByVal wb As Workbook): End Sub

'@sub-title Import keys and metadata from another password handler instance.
Public Sub ImportFrom(ByVal source As IPasswords)
End Sub

'@fun-title Clone the password worksheet into another sheet and returns its handler.
Public Function CloneToWorksheet(ByVal targetSheet As Worksheet) As IPasswords
End Function

'@fun-title Clone the password worksheet into another workbook and returns its handler
'@details
'this will move the password Object focus to the added workbook. All the new
'worksheets search (with worksheet name) are done on the added workbook. The
'password is transfered to the new workbook (keys table and protection table),
'and the new workbook has a debug mode exit handler.
Public Function CloneToWorkbook(ByVal wb As Workbook) As IPasswords
End Function
