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


'@section Enumerations
'===============================================================================
'@description
'Project-wide enumerations for diagnostic scopes, value selectors, and
'error codes. These lived in the removed IChecking interface; they are
're-homed on the Checking class so it is their single source of truth.
'Depending on Checking is enough to resolve them everywhere.

'Severity levels for diagnostic entries.
Public Enum CheckingScope
  checkingError = 1
  checkingWarning = 2
  checkingInfo = 3
  checkingNote = 4
  checkingSuccess = 5
End Enum

'Selectors for retrieving entry properties.
Public Enum CheckingValue
  checkingLabel = 1
  checkingType = 2
  checkingColor = 3
End Enum

'Application-wide error codes used with Err.Raise.
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
  sealed As Boolean
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
'@return Checking. A ready-to-use Checking instance.
'@depends BetterArray
'@export
Public Function Create(ByVal titleName As String, _
                       Optional ByVal subtitleName As String = vbNullString) As _
                      Checking
  Dim newChecking As Checking
  Set newChecking = New Checking

  newChecking.CheckingTitle = titleName
  newChecking.CheckingSubTitle = subtitleName
  newChecking.Seal
  Set Create = newChecking
End Function

'@section PublicAccessors
'===============================================================================
'@description Properties that expose internal state and the object pointer.

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
  GuardNotSealed "CheckingTitle"
  this.title = titleName
End Property

'@label:checking-subtitle-set
'@prop-title Assign the checking subtitle
'@param subtitleName String. Subtitle to assign.
Public Property Let CheckingSubTitle(ByVal subtitleName As String)
  GuardNotSealed "CheckingSubTitle"
  this.subtitle = subtitleName
End Property

'@section Elements
'===============================================================================
'@description Read-only metadata properties.

'@label:name
'@prop-title Name of the checking object
'@details
'Returns the checking title as the object name. Used for identification
'in output and presentation contexts.
'@return String. The checking title.
'@export
Public Property Get Name() As String
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
Public Property Get Heading(Optional ByVal subtitle As Boolean = False) As String
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
Public Property Get Length() As Long
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
Public Property Get ListOfKeys() As BetterArray

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
'
'The stored table is read directly. `ListOfKeys` hands out a clone, and `Add`
'asks this question on every call, so a build that files 500 entries used to
'copy an array averaging 250 items 500 times before recording anything.
'@param keyName String. Key to search for.
'@return Boolean. True when the key is found.
'@export
Public Function KeyExists(ByVal keyName As String) As Boolean

  If (keyName = vbNullString) Then Exit Function
  If this.keysTable Is Nothing Then Exit Function

  KeyExists = this.keysTable.Includes(keyName)
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
 Public Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = 1) As String

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
'@return Checking. Independent clone.
'@export
Public Function Clone() As Checking
  Dim cloneInstance As Checking
  Dim idx As Long
  Dim lowerBound As Long
  Dim upperBound As Long

  Set cloneInstance = New Checking
  cloneInstance.CheckingTitle = this.title
  cloneInstance.CheckingSubTitle = this.subtitle
  cloneInstance.Seal

  If this.nbElmts <= 0 Then
    Set Clone = cloneInstance
    Exit Function
  End If

  If this.keysTable Is Nothing Then
    Set Clone = cloneInstance
    Exit Function
  End If

  lowerBound = this.keysTable.LowerBound
  upperBound = this.keysTable.UpperBound

  For idx = lowerBound To upperBound
    cloneInstance.Add CStr(this.keysTable.Item(idx)), _
                     CStr(this.labelsTable.Item(idx)), _
                     CByte(this.scopesTable.Item(idx))
  Next

  Set Clone = cloneInstance
End Function

'@label:append
'@sub-title Append entries from another checking to the current one
'@details
'Iterates through the supplied Checking entries and adds each one
'to the current checking. Silently exits when the supplied checking
'is Nothing or empty.
'
'WHAT HAPPENS TO A KEY BOTH SIDES HOLD
'-------------------------------------------------------------------------------
'It is numbered and the entry is kept. `Add` is the strict half of the contract
'and still raises on a duplicate; this is the merge, and every caller in the
'tree uses it as one. Two classes that both key their entries with a counter
'from 0 -- `Formulas` and `FormulaCondition` do -- used to stop the caller dead
'the second time their stores met.
'
'Owner decision, 2026-07-31.
'@param checks Checking. Source checking to merge from.
'@export
Public Sub Append(ByVal checks As Checking)
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

    Add FreeKey(keyName), _
      checks.ValueOf(keyName, checkingLabel), _
      keyScope
  Next idx
End Sub

'@label:free-key
'@sub-title Answer a key this checking does not hold yet
'@details
'The root is numbered until it is free. Used by `Append` alone, so a direct
'`Add` keeps raising on a duplicate.
'@param keyRoot String. The key to start from.
'@return String. A key that is free.
Private Function FreeKey(ByVal keyRoot As String) As String
  Dim candidate As String
  Dim suffix As Long

  candidate = keyRoot
  Do While KeyExists(candidate)
    suffix = suffix + 1
    candidate = keyRoot & " #" & CStr(suffix)
  Loop

  FreeKey = candidate
End Function

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
