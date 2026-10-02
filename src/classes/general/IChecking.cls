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
'@Folder("General")
'@ModuleDescription("Interface of setup checkings")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'@section Enumerations
'===============================================================================
'@description Project-wide enumerations for diagnostic scopes, values, and error codes.

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

Option Explicit

'@class IChecking
'@description
'Defines the public contract for the Checking class, exposing entry
'management (add, clone, append), key lookup, and property retrieval
'for diagnostic messages tagged with severity scopes. Also hosts
'project-wide enumerations for CheckingScope, CheckingValue, and
'ProjectError.


'@section Elements
'===============================================================================
'@description Read-only metadata properties.

'@jump:name
'@prop-title Name of the checking object
'@details
'Returns the checking title as the object name.
'@return String. The checking title.
Public Property Get Name() As String
End Property

'@jump:heading
'@prop-title Title or subtitle of the checking object
'@details
'Returns the subtitle when the flag is True, otherwise returns the title.
'@param subtitle Optional Boolean. When True, returns the subtitle. Defaults to False.
'@return String. The title or subtitle.
Public Property Get Heading(Optional ByVal subtitle As Boolean = False) As String
End Property

'@jump:length
'@prop-title Total number of entries in the checking
'@details
'Returns the count of diagnostic messages stored.
'@return Long. Number of entries.
Public Property Get Length() As Long
End Property


'@section Keys
'===============================================================================
'@description Key enumeration and lookup.

'@jump:list-of-keys
'@prop-title Get the list of keys in the checking
'@details
'Returns a clone of the internal keys array. Returns an empty
'BetterArray when no entries have been added.
'@return BetterArray. Clone of stored keys.
Public Property Get ListOfKeys() As BetterArray
End Property

'@jump:key-exists
'@prop-title Test whether a key is present in the checking
'@details
'Returns True when the specified key exists in the entries.
'@param keyName String. Key to search for.
'@return Boolean. True when the key is found.
Public Property Get KeyExists(ByVal keyName As String) As Boolean
End Property


'@section InputOutput
'===============================================================================
'@description Adding, retrieving, cloning, and appending diagnostic entries.

'@jump:add
'@sub-title Add a new entry to the checking
'@details
'Adds a key, label, and scope to the checking. Keys must be unique.
'@param keyName String. Unique identifier for the entry.
'@param label String. Diagnostic message for the entry.
'@param scope Optional Byte. Severity scope. Defaults to checkingNote.
Public Sub Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = checkingNote)
End Sub

'@jump:value-of
'@prop-title Retrieve a property of one specific key
'@details
'Returns the label, type descriptor, or colour string depending on
'the retrievedValue selector.
'@param keyName String. Key from which to retrieve values.
'@param retrievedValue Optional Byte. CheckingValue enum selector. Defaults to checkingLabel.
'@return String. The requested property value.
Public Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = 1) As String
End Property

'@jump:clone
'@sub-title Return an independent copy of the current checking object
'@details
'Creates a deep copy with the same title, subtitle, and all entries.
'@return IChecking. Independent clone.
Public Function Clone() As IChecking
End Function

'@jump:append
'@sub-title Add entries from another checking to the current one
'@details
'Merges all entries from the supplied checking into this one.
'@param checks IChecking. Source checking to merge from.
Public Sub Append(ByVal checks As IChecking)
End Sub
