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

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Save warnings, information and errors on objects")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit
Implements IChecking

'@class Checking
'@description
'Stores a collection of diagnostic messages, each tagged with a unique
'key, a label, and a severity scope (Error, Warning, Info, Note, or
'Success). Checking objects serve as lightweight log containers that
'other classes populate during validation or build processes. The entries
'are stored in three parallel BetterArray tables (keys, labels, scopes).
'Presentation to a worksheet is handled by the separate CheckingOutput
'class.
'@depends BetterArray
'@version 1.0 (2026-02-09)
'@author OutbreakTools


'@section State
'===============================================================================
'@description Backing UDT and class constants.

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

'@section Factory
'===============================================================================
'@description Entry point for creating Checking instances.

'@label:create
'@sub-title Create a Checking instance for logging diagnostic messages
'@details
'Checking objects are collections of messages tagged with a scope
'(Error, Warning, Info, Note, or Success) for logging purposes. The
'entries are stored in three parallel BetterArray tables (keys, labels,
'scopes). Checkings must have a title identifying the topic. Optionally,
'a subtitle provides further context. Titles and subtitles are used
'when printing the messages to a worksheet via CheckingOutput.
'@param titleName String. Title of the checkings topic.
'@param subtitleName Optional String. Subtitle of the checkings. Defaults to vbNullString.
'@return IChecking. A ready-to-use Checking instance.
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

'@section PublicAccessors
'===============================================================================
'@description Properties that expose internal state and the interface pointer.

'@label:self
'@prop-title Current object instance
'@details
'Returns the current instance cast to its interface. Used internally
'by the factory method to return an IChecking reference.
'@return IChecking. Reference to the current object.
Public Property Get Self() As IChecking
  Set Self = Me
End Property

'@label:checking-title
'@prop-title Title of the checking topic
'@details
'Returns the title string assigned during creation or via the setter.
'@return String. Title of the checkings.
Public Property Get CheckingTitle() As String
  CheckingTitle = this.title
End Property

'@label:checking-subtitle
'@prop-title Subtitle of the checking topic
'@details
'Returns the subtitle string assigned during creation or via the setter.
'@return String. Subtitle of the checkings.
Public Property Get CheckingSubTitle() As String
  CheckingSubTitle = this.subtitle
End Property

'@label:checking-title-set
'@prop-title Assign the checking title
'@param titleName String. Title to assign.
Public Property Let CheckingTitle(ByVal titleName As String)
  this.title = titleName
End Property

'@label:checking-subtitle-set
'@prop-title Assign the checking subtitle
'@param subtitleName String. Subtitle to assign.
Public Property Let CheckingSubTitle(ByVal subtitleName As String)
  this.subtitle = subtitleName
End Property

'@section Elements
'===============================================================================
'@description Read-only metadata properties exposed through the interface.

'@label:name
'@prop-title Name of the checking object
'@details
'Returns the checking title as the object name. Used for identification
'in output and presentation contexts.
'@return String. The checking title.
'@export
Private Property Get Name() As String
  Name = CheckingTitle()
End Property

'@label:heading
'@prop-title Title or subtitle of the checking object
'@details
'Returns the subtitle when the subtitle flag is True, otherwise returns
'the title. Used by CheckingOutput for heading display.
'@param subtitle Optional Boolean. When True, returns the subtitle. Defaults to False.
'@return String. The title or subtitle.
'@export
Private Property Get Heading(Optional ByVal subtitle As Boolean = False) As String
  If subtitle Then
    Heading = CheckingSubTitle()
    Exit Property
  End If
  Heading = CheckingTitle()
End Property

'@label:length
'@prop-title Total number of entries in the checking
'@details
'Returns the count of diagnostic messages stored in the parallel arrays.
'@return Long. Number of entries.
'@export
Private Property Get Length() As Long
  Length = this.nbElmts
End Property

'@section Keys
'===============================================================================
'@description Key lookup and enumeration.

'@label:list-of-keys
'@prop-title Get the list of keys in the checking
'@details
'Returns a clone of the internal keys BetterArray. When the keys table
'is uninitialised, creates an empty BetterArray with LowerBound 1.
'@return BetterArray. Clone of stored keys.
'@export
Private Property Get ListOfKeys() As BetterArray

  'If the keysTable is unitialised, initialize it
  If this.keysTable Is Nothing Then
    Set this.keysTable = New BetterArray
    this.keysTable.lowerBound = 1
  End If
  Set ListOfKeys = this.keysTable.Clone()
End Property

'@label:key-exists
'@sub-title Test whether a key is present in the checking
'@details
'Searches the keys BetterArray for the specified key name. Returns
'False immediately for empty strings.
'@param keyName String. Key to search for.
'@return Boolean. True when the key is found.
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

'@section InputOutput
'===============================================================================
'@description Adding, retrieving, cloning, and appending diagnostic entries.

'@label:add
'@sub-title Add a new entry to the checking
'@details
'Adds a key, the corresponding label, and the scope to the checking.
'Keys must be unique; duplicate keys raise an error. On first call,
'initialises the three parallel BetterArray tables.
'@param keyName String. Unique identifier for the entry.
'@param label String. Diagnostic message for the entry.
'@param scope Optional Byte. Severity scope (checkingError, checkingWarning, etc.). Defaults to checkingNote.
'@throws ProjectError.ElementShouldNotExists When the key already exists.
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

'@label:value-of
'@prop-title Retrieve a property of one specific key
'@details
'Returns the label, the type descriptor (with icon), or the colour
'string for the specified key depending on the retrievedValue selector.
'Raises an error when the key does not exist.
'@param keyName String. Key from which to retrieve values.
'@param retrievedValue Optional Byte. CheckingValue enum selector. Defaults to checkingLabel.
'@return String. The requested property value.
'@throws ProjectError.ElementShouldExists When the key is not found.
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
  actualScope = this.scopesTable.Item(indexValue)

  Select Case retrievedValue

    Case checkingLabel
      returnedValue = this.labelsTable.Item(indexValue)
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

'@label:clone
'@sub-title Produce a deep copy of the current checking entries
'@details
'Creates a new Checking instance with the same title, subtitle, and
'logged entries. Required so that flushing a checking into another
'container does not keep a live reference to the current object.
'@return IChecking. Independent clone.
'@export
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

  lowerBound = this.keysTable.LowerBound
  upperBound = this.keysTable.UpperBound

  For idx = lowerBound To upperBound
    cloneInstance.Add CStr(this.keysTable.Item(idx)), _
                     CStr(this.labelsTable.Item(idx)), _
                     CByte(this.scopesTable.Item(idx))
  Next

  Set Clone = cloneInstance.Self()
End Function

'@label:append
'@sub-title Append entries from another checking to the current one
'@details
'Iterates through the supplied IChecking entries and adds each one
'to the current checking. Silently exits when the supplied checking
'is Nothing or empty.
'@param checks IChecking. Source checking to merge from.
'@export
Private Sub Append(ByVal checks As IChecking)
  Dim idx As Long
  Dim keyTable As BetterArray
  Dim keyName As String
  Dim keyScope As Byte
  Dim keyTypeLabel As String

  If (checks Is Nothing) Then Exit Sub
  If checks.Length() <= 0 Then Exit Sub

  Set keyTable = checks.ListOfKeys()

  If (keyTable Is Nothing) Then Exit Sub

  For idx = keyTable.LowerBound To keyTable.UpperBound
    keyName = CStr(keyTable.Item(idx))
    keyTypeLabel = checks.ValueOf(keyName, checkingType)
    keyScope = ScopeFromTypeLabel(keyTypeLabel)

    Add keyName, _
      checks.ValueOf(keyName, checkingLabel), _
      keyScope
  Next idx
End Sub

'@section Helpers
'===============================================================================
'@description Private utility methods.

'@label:scope-from-type-label
'@sub-title Convert a type label string back to its CheckingScope byte value
'@details
'Parses the human-readable type label (e.g. "Error", "Warning") and
'returns the corresponding CheckingScope enum value. Defaults to
'checkingNote when no match is found.
'@param typeLabel String. Human-readable type label.
'@return Byte. CheckingScope value.
Private Function ScopeFromTypeLabel(ByVal typeLabel As String) As Byte
  Dim cleaned As String

  cleaned = Trim$(typeLabel)

  If cleaned = vbNullString Then
    ScopeFromTypeLabel = checkingNote
    Exit Function
  End If

  If InStr(1, cleaned, "Error", vbTextCompare) > 0 Then
    ScopeFromTypeLabel = checkingError
  ElseIf InStr(1, cleaned, "Warning", vbTextCompare) > 0 Then
    ScopeFromTypeLabel = checkingWarning
  ElseIf InStr(1, cleaned, "Info", vbTextCompare) > 0 Then
    ScopeFromTypeLabel = checkingInfo
  ElseIf InStr(1, cleaned, "Success", vbTextCompare) > 0 Then
    ScopeFromTypeLabel = checkingSuccess
  Else
    ScopeFromTypeLabel = checkingNote
  End If
End Function

'@section ErrorHandling
'===============================================================================
'@description Centralised error-raising pattern.

'@label:throw-error
'@sub-title Raise a VBA error with the class signature
'@details
'Wrapper around Err.Raise that standardises the source to CLASSNAME
'for consistent stack traces.
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error message : " & Chr(34) & errorMessage
End Sub

'@section Interface
'===============================================================================
'@description Delegated members satisfying the IChecking contract.

Private Property Get IChecking_Name() As String
  IChecking_Name = Name()
End Property

Private Property Get IChecking_Heading(Optional ByVal subtitle As Boolean = False) As String
  IChecking_Heading = Heading(subtitle)
End Property

Private Property Get IChecking_Length() As Long
  IChecking_Length = Length()
End Property

Private Property Get IChecking_ListOfKeys() As BetterArray
  Set IChecking_ListOfKeys = ListOfKeys()
End Property

Private Property Get IChecking_KeyExists(ByVal keyName As String) As Boolean
  IChecking_KeyExists = KeyExists(keyName)
End Property

Private Sub IChecking_Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = checkingNote)
    Add keyName, label, scope
End Sub

Private Property Get IChecking_ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = 1) As String
  IChecking_ValueOf = ValueOf(keyName, retrievedValue)
End Property

Private Function IChecking_Clone() As IChecking
  Set IChecking_Clone = Clone()
End Function

Private Sub IChecking_Append(ByVal checks As IChecking)
  Append checks
End Sub
