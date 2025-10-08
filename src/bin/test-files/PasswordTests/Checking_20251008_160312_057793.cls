VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Save warnings, information and errors on objects"


'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Save warnings, information and errors on objects")

Option Explicit
Implements IChecking

Private Const CLASSNAME As String = "Checking"

Private Type TChecking
  title As String
  subtitle As String
  keysTable As BetterArray
  labelsTable As BetterArray
  scopesTable As BetterArray
  nbElmts As Long
End Type

Private this As TChecking

'@label:Create
'@prop-title Create a checking object
'@details
'Checking objects are a collection of messages tagged with some scope (Error, Warning, Notes)
'for logging purposes. The checkings are stored in three betterarray tables
'(one for label, one for scopes, and one for keys).
'Checkings must have a name, which is the title of the messages. One can create
' checkings on specific thematics. Optionally, checkings could have subtitles.
'Titles and subtitles are used for writing and presenting the checking messages
'in a provided worksheets. Printing to the outside world ( a worksheet) is done with another class that
'relies on the Checkings.
'
'@param titleName A String, the title of the checkings
'@param subtitleName Optional, String. Subtitle of the checkings. Default is vbNullString
'@return A IChecking Object
'@depends BetterArray
'@export
Public Function Create(ByVal titleName As String, _
                       Optional ByVal subtitleName As String = vbNullString) As _
                      IChecking
  With New Checking
    .CheckingTitle = titleName
    .CheckingSubTitle = subtitleName
    Set Create = .Self()
  End With
End Function

'@label:elements
'@section Checking Elements
'===============================================================================

Public Property Get Self() As IChecking
  Set Self = Me
End Property

'title and subTitle
Public Property Get CheckingTitle() As String
  CheckingTitle = this.title
End Property

Public Property Get CheckingSubTitle() As String
  CheckingSubTitle = this.subtitle
End Property

Public Property Let CheckingTitle(ByVal titleName As String)
  this.title = titleName
End Property

Public Property Let CheckingSubTitle(ByVal subtitleName As String)
  this.subtitle = subtitleName
End Property

'@label:Name
'@prop-title Name of the checking object
'@export
'@return A String
Private Property Get Name() As String
  Name = CheckingTitle()
End Property

'@label:Heading
'@prop-title Title or subtitle of a checking object
'@param subTitle Boolean whether to return to subtitle instead of the title. Default is False
'@export
'@return A String
Private Property Get Heading(Optional ByVal subtitle As Boolean = False) As String
  If subtitle Then
    Heading = CheckingSubTitle()
    Exit Property
  End If
  Heading = CheckingTitle()
End Property

'@label:Length
'@prop-title Total number of elements in the checking
'@export
'@return A Long
Private Property Get Length() As Long
  Length = this.nbElmts
End Property

'@label:keys
'@section Keys
'===============================================================================

'@label:ListOfKeys
'@prop-title Get the list of keys in the checking
'@return A BetterArray
Private Property Get ListOfKeys() As BetterArray

  'If the keysTable is unitialised, initialize it
  If this.keysTable Is Nothing Then
    Set this.keysTable = New BetterArray
    this.keysTable.lowerBound = 1
  End If
  Set ListOfKeys = this.keysTable.Clone()
End Property

'@label:keyExists
'@prop-title Test if a key is in the checking
'@param keyName A String, the key to search
'@return A Boolean
'@export
Private Function KeyExists(ByVal keyName As String) As Boolean
  Dim keyTabl As BetterArray

  If (keyName = vbNullString) Then
    KeyExists = False
    Exit Function
  End If

  Set keyTabl = ListOfKeys()
  KeyExists = keyTabl.Includes(keyName)
End Function

'@label:io
'@section Input / Output
'===============================================================================

'@label:Add
'@prop-title Add new values to the checking
'@details
'Add a key, the corresponding label and the scope of the label to a checking
'object. Keys should be unique. In case keys are not, an error is raised
'@param keyName A String
'@param label A String, the message corresponding to the key
'@param scope Optional, Byte. The checking scope (Error, Warning, Info)
'@export
Public Sub Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = checkingNote)

  If KeyExists(keyName) Then
    ThrowError ProjectError.ElementShouldNotExists, _
               keyName & " already exists, keys should be unique."
  End If

  'Initialize the tables
  If (this.labelsTable Is Nothing) And (this.scopesTable Is Nothing) Then
    Set this.labelsTable = New BetterArray
    Set this.scopesTable = New BetterArray
    Set this.keysTable = New BetterArray

    this.labelsTable.lowerBound = 1
    this.scopesTable.lowerBound = 1
    this.keysTable.lowerBound = 1
  End If

  'Push the three values
  this.keysTable.Push keyName
  this.scopesTable.Push scope
  this.labelsTable.Push label
  this.nbElmts = this.nbElmts + 1
End Sub

'@label:ValueOf
'@prop-title Retrieve a property of one specific key
'@details
'Get either the label or the type of one specific key in the checkings
'@param keyName A String the key from which to retrieve values
'@param retrievedValue A Byte, checkingValue Enum. Value to retrieve. Default is checkingLabel.
'@return A String
'@export
 Private Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = 1) As String

  Dim indexValue As Long
  Dim returnedValue As String
  Dim actualScope As Byte

  If (Not KeyExists(keyName)) Then
    ThrowError ProjectError.ElementShouldExists, _
               keyName & " does not exists in this checking"
  End If

  indexValue = this.keysTable.IndexOf(keyName)
  actualScope = this.scopesTable.item(indexValue)

  Select Case retrievedValue

    Case checkingLabel
      returnedValue = this.labelsTable.item(indexValue)
    Case checkingType
      Select Case actualScope
        Case checkingNote
          returnedValue = ChrW(9998) & " Note"
        Case checkingError
          returnedValue = ChrW(10060) & " Error"
        Case checkingWarning
          returnedValue = ChrW(9888) & " Warning"
        Case checkingInfo
          returnedValue = ChrW(8505) & " Info"
        Case checkingSuccess
          returnedValue = ChrW(10004) & " Success"
      End Select
    Case checkingColor
      Select Case actualScope
        Case checkingNote
          returnedValue = "purple"
        Case checkingError
          returnedValue = "red"
        Case checkingWarning
          returnedValue = "orange"
        Case checkingInfo
          returnedValue = "grey"
        Case checkingSuccess
          returnedValue = "green"
      End Select
  End Select
  ValueOf = returnedValue
End Property

'@label:Clone
'@prop-title Produce a deep copy of the current checking entries
'@details
'Creates a new checking instance with the same title, subtitle, and logged
'entries. This is required so that flushing a checking into another
'BetterArray does not keep a live reference to the current object, which would
'be cleared afterwards.
'@return IChecking independent clone
Private Function Clone() As IChecking
  Dim cloneInstance As Checking
  Dim idx As Long
  Dim lowerBound As Long
  Dim upperBound As Long

  Set cloneInstance = New Checking
  cloneInstance.CheckingTitle = this.title
  cloneInstance.CheckingSubTitle = this.subtitle

  If this.nbElmts <= 0 Then
    Set Clone = cloneInstance.Self
    Exit Function
  End If

  If this.keysTable Is Nothing Then
    Set Clone = cloneInstance.Self
    Exit Function
  End If

  lowerBound = this.keysTable.lowerBound
  upperBound = this.keysTable.upperBound

  For idx = lowerBound To upperBound
    cloneInstance.Add CStr(this.keysTable.item(idx)), _
                     CStr(this.labelsTable.item(idx)), _
                     CByte(this.scopesTable.item(idx))
  Next idx

  Set Clone = cloneInstance.Self()
End Function


'Errors and checkings -----------------------------------------------------------
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error message : " & Chr(34) & errorMessage
End Sub

'@jump:Create
'@section Interface
'===============================================================================

'@jump:elements
'Checking elements
'===============================================================================

'name of the checking object
'@jump:Name
Private Property Get IChecking_Name() As String
  IChecking_Name = Name()
End Property

'Title or subtitle of the checking object
Private Property Get IChecking_Heading(Optional ByVal subtitle As Boolean = False) As String
  IChecking_Heading = Heading(subtitle)
End Property

'Length of the checkings
Private Property Get IChecking_Length() As Long
  IChecking_Length = Length()
End Property

'@jump:keys
'===============================================================================

'@jump:ListOfKeys
Private Property Get IChecking_ListOfKeys() As BetterArray
  Set IChecking_ListOfKeys = ListOfKeys()
End Property

'@jump:keyExists
Private Property Get IChecking_KeyExists(ByVal keyName As String) As Boolean
  IChecking_KeyExists = KeyExists(keyName)
End Property

'@jump:io
'Input/Ouptut
'===============================================================================

'Exposed methods
Private Sub IChecking_Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = checkingNote)
    Add keyName, label, scope
End Sub


'You can retrieve the label, the scope or the color
Private Property Get IChecking_ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = 1) As String
  IChecking_ValueOf = ValueOf(keyName, retrievedValue)
End Property

Private Function IChecking_Clone() As IChecking
  Set IChecking_Clone = Clone()
End Function
