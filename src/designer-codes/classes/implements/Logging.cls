VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Logging informations for the designer"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("Logging informations for the designer")

Option Explicit
Implements ILogging

Private Const CLASSNAME As String = "Logging"

Private Type TLogging
  title As String
  subTitle As String
  keysTable As BetterArray
  labelsTable As BetterArray
  scopesTable As BetterArray
End Type

Private this As TLogging

'Exposed methods
Public Function Create(ByVal titleName As String, _
                       Optional ByVal subtitleName As String = vbNullString) As ILogging
  If titleName = vbNullString Then ThrowError "Loggings should have a title"
  With New Logging
    .LoggingTitle = titleName
    .LoggingSubTitle = subtitleName
    Set Create = .Self()
  End With
End Function

Public Property Get Self() As ILogging
  Set Self = Me
End Property

'title and subTitle
Public Property Get LoggingTitle() As String
  LoggingTitle = this.title
End Property

Public Property Get LoggingSubTitle() As String
  LoggingSubTitle = this.subTitle
End Property

Public Property Let LoggingTitle(ByVal titleName As String)
  this.title = titleName
End Property

Public Property Let LoggingSubTitle(ByVal subtitleName As String)
  this.subTitle = subtitleName
End Property

'Heading Property
Private Property Get Heading(Optional ByVal subTitle As Boolean) As String
  If subTitle Then
    Heading = LoggingSubTitle()
    Exit Property
  End If
  Heading = LoggingTitle()
End Property


'Get the list of keys
Private Property Get ListOfKeys() As BetterArray

  'If the keysTable is unitialised, initialize it
  If this.keysTable Is Nothing Then
    Set this.keysTable = New BetterArray
    this.keysTable.LowerBound = 1
  End If
  Set ListOfKeys = this.keysTable.Clone()
End Property

'Test if a key already exists
Private Function keyExists(ByVal keyName As String) As Boolean
  Dim keyTabl As BetterArray

  If (keyName = vbNullString) Then
    keyExists = False
    Exit Function
  End If

  Set keyTabl = ListOfKeys()
  keyExists = keyTabl.Includes(keyName)
End Function

'Add new values to the Logging

Private Sub Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = LoggingNote)

  If keyExists(keyName) Then ThrowError keyName & " already exists, keys should be unique."

  'Initialize the tables
  If (this.labelsTable Is Nothing) And (this.scopesTable Is Nothing) Then
    Set this.labelsTable = New BetterArray
    Set this.scopesTable = New BetterArray
    Set this.keysTable = New BetterArray

    this.labelsTable.LowerBound = 1
    this.scopesTable.LowerBound = 1
    this.keysTable.LowerBound = 1
  End If

  'Push the three values
  this.keysTable.Push keyName
  this.scopesTable.Push scope
  this.labelsTable.Push label
End Sub

'Get either the label or the type of one key in the Loggings
'The Logging value could be either the Logging label,
'(the text) or the Logging color or the LoggingType (Note, warning, info Or Error)
Private Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As LoggingValue = 1) As String
  Dim indexValue As Long
  Dim returnedValue As String
  Dim actualScope As LoggingScope

  If (Not keyExists(keyName)) Then ThrowError keyName & " does not exists in this Logging"

  indexValue = this.keysTable.IndexOf(keyName)
  actualScope = this.scopesTable.Item(indexValue)

  Select Case retrievedValue

    Case LoggingLabel
      returnedValue = this.labelsTable.Item(indexValue)
    Case LoggingType
      Select Case actualScope
        Case LoggingNote
          returnedValue = "Note"
        Case LoggingError
          returnedValue = "Error"
        Case LoggingWarning
          returnedValue = "Warning"
        Case LoggingInfo
          returnedValue = "Info"
      End Select
    Case LoggingColor
      Select Case actualScope
        Case LoggingNote
          returnedValue = "blue"
        Case LoggingError
          returnedValue = "red"
        Case LoggingWarning
          returnedValue = "orange"
        Case LoggingInfo
          returnedValue = "grey"
      End Select
  End Select
  ValueOf = returnedValue
End Property


'Errors and Loggings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub


'Exposed methods
Private Sub ILogging_Add(ByVal keyName As String, ByVal label As String, _
                Optional ByVal scope As Byte = LoggingNote)
    Add keyName, label, scope
End Sub
Private Property Get ILogging_ListOfKeys() As BetterArray
  Set ILogging_ListOfKeys = ListOfKeys()
End Property

'Heading could be subtitle or title.
Private Property Get ILogging_Heading(Optional ByVal subTitle As Boolean = False) As String
  ILogging_Heading = Heading(subTitle)
End Property

'You can retrieve the label, the scope or the color
Private Property Get ILogging_ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As Byte = LoggingLabel) As String
  ILogging_ValueOf = ValueOf(keyName, retrievedValue)
End Property
