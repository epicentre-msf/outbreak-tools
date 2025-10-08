VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of setup checkings"

'@Interface
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("General")
'@ModuleDescription("Interface of setup checkings")

'Scopes for different checks
Public Enum CheckingScope
  checkingError = 1
  checkingWarning = 2
  checkingInfo = 3
  checkingNote = 4
  checkingSuccess = 5
End Enum

'Checking Values
Public Enum CheckingValue
  checkingLabel = 1
  checkingType = 2
  checkingColor = 3
End Enum

'Project Errors
Public Enum ProjectError
  ErrorNone = 0
  InvalidArgument = 1001
  ObjectNotInitialized = 1002
  ErrorUnexpectedState = 1004
  ElementShouldExists = 1005
  ElementShouldNotExists = 1006
  ElementNotFound = 1007
  SomethingWentWrong = 1008
End Enum

Option Explicit

'@jump:Create
'===============================================================================

'@jump:elements
'@section Elements
'===============================================================================
'@jump:Name
Public Property Get Name() As String
End Property

'@jump:Heading
'Title or Subtitle of checkings
Public Property Get Heading(Optional ByVal subtitle As Boolean = False) As String
End Property

'@jump:keys
'@section Keys
'===============================================================================

'@jump:ListOfKeys
Public Property Get ListOfKeys() As BetterArray
End Property

'@jump:keyExists
Public Property Get KeyExists(ByVal keyName As String) As Boolean
End Property

'@jump:Length
Public Property Get Length() As Long
End Property

'@jump:io
'@section Input / Output
'===============================================================================

'@jump:Add
'Add new values to the checking
Public Sub Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = checkingNote)
End Sub

'@jump:ValueOf
'Retrieve a property of one specific key
Public Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = 1) As String
End Property

'@jump:Clone
'Return an independent copy of the current checking object
Public Function Clone() As IChecking
End Function
