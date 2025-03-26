VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of setup Loggings"


'@Folder("Checks")
'@ModuleDescription("Interface of setup Loggings")

Public Enum LoggingScope
  LoggingError = 1
  LoggingWarning = 2
  LoggingInfo = 3
  LoggingNote = 4
End Enum

Public Enum LoggingValue
  LoggingLabel = 1
  LoggingType = 2
  LoggingColor = 3
End Enum

Option Explicit


'Exposed methods
Public Sub Add(ByVal keyName As String, ByVal LABEL As String, _
                Optional ByVal scope As Byte = LoggingNote)
End Sub
Public Property Get ListOfKeys() As BetterArray
End Property
'Heading could be subtitle or title.
Public Property Get Heading(Optional ByVal subTitle As Boolean = False) As String
End Property
'You can retrieve the label, the scope or the color
Public Property Get ValueOf(ByVal keyName As String, _
                            Optional ByVal retrievedValue As Byte = LoggingLabel) As String
End Property
