VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Checking of the setup"

'@PredeclaredId
'@Folder("Checks")
'@ModuleDescription("Checking of the setup")

Option Explicit
Implements IChecking

Private Const CLASSNAME As String = "Checking"

Private Type TChecking
  title As String
  subTitle As String
  keysTable As BetterArray
  labelsTable As BetterArray
  scopesTable As BetterArray
End Type

Private this As TChecking

'Exposed methods
Public Function Create(ByVal titleName As String, _
                       Optional ByVal subtitleName As String = vbNullString) As IChecking
  If titleName = vbNullString Then ThrowError "Checkings should have a title"
  With New Checking
    .CheckingTitle = titleName
    .CheckingSubTitle = subtitleName
    Set Create = .Self()
  End With
End Function

Public Property Get Self() As IChecking
  Set Self = Me
End Property

'title and subTitle
Public Property Get CheckingTitle() As String
  CheckingTitle = this.title
End Property

Public Property Get CheckingSubTitle() As String
  CheckingSubTitle = this.subTitle
End Property

Public Property Let CheckingTitle(ByVal titleName As String)
  this.title = titleName
End Property

Public Property Let CheckingSubTitle(ByVal subtitleName As String)
  this.subTitle = subtitleName
End Property

'Heading Property
Private Property Get Heading(Optional ByVal subTitle As Boolean) As String
  If subTitle Then
    Heading = CheckingSubTitle()
    Exit Property
  End If
  Heading = CheckingTitle()
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

'Add new values to the checking

Private Sub Add(ByVal keyName As String, ByVal LABEL As String, _
                Optional ByVal scope As CheckingScope = checkingNote)

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
  this.labelsTable.Push LABEL
End Sub

'Get either the label or the type of one key in the checkings
'The checking value could be either the checking label,
'(the text) or the checking color or the checkingType (Note, warning, info Or Error)
Private Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As CheckingValue = 1) As String
  Dim indexValue As Long
  Dim returnedValue As String
  Dim actualScope As CheckingScope

  If (Not keyExists(keyName)) Then ThrowError keyName & " does not exists in this checking"

  indexValue = this.keysTable.IndexOf(keyName)
  actualScope = this.scopesTable.Item(indexValue)

  Select Case retrievedValue

    Case checkingLabel
      returnedValue = this.labelsTable.Item(indexValue)
    Case checkingType
      Select Case actualScope
        Case checkingNote
          returnedValue = "Note"
        Case checkingError
          returnedValue = "Error"
        Case checkingWarning
          returnedValue = "Warning"
        Case checkingInfo
          returnedValue = "Info"
      End Select
    Case checkingColor
      Select Case actualScope
        Case checkingNote
          returnedValue = "blue"
        Case checkingError
          returnedValue = "red"
        Case checkingWarning
          returnedValue = "orange"
        Case checkingInfo
          returnedValue = "grey"
      End Select
  End Select
  ValueOf = returnedValue
End Property


'Errors and checkings
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub


'Exposed methods
Private Sub IChecking_Add(ByVal keyName As String, ByVal LABEL As String, _
                Optional ByVal scope As CheckingScope = checkingNote)
    Add keyName, LABEL, scope
End Sub
Private Property Get IChecking_ListOfKeys() As BetterArray
  Set IChecking_ListOfKeys = ListOfKeys()
End Property

'Heading could be subtitle or title.
Private Property Get IChecking_Heading(Optional ByVal subTitle As Boolean = False) As String
  IChecking_Heading = Heading(subTitle)
End Property

'You can retrieve the label, the scope or the color
Private Property Get IChecking_ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As CheckingValue = 1) As String
  IChecking_ValueOf = ValueOf(keyName, retrievedValue)
End Property
