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


'@Folder("Checks")
'@ModuleDescription("Interface of setup checkings")

Public Enum CheckingScope
  checkingError = 1
  checkingWarning = 2
  checkingInfo = 3
  checkingNote = 4
End Enum

Public Enum CheckingValue
  checkingLabel = 1
  checkingType = 2
  checkingColor = 3
End Enum

Option Explicit


'Exposed methods
Public Sub Add(ByVal keyName As String, ByVal LABEL As String, _
                Optional ByVal scope As CheckingScope = checkingNote)
End Sub
Public Property Get ListOfKeys() As BetterArray: End Property
'Heading could be subtitle or title.
Public Property Get Heading(Optional ByVal subTitle As Boolean = False) As String: End Property
'You can retrieve the label, the scope or the color
Public Property Get ValueOf(ByVal keyName As String, Optional ByVal retrievedValue As CheckingValue = 1) As String: End Property
