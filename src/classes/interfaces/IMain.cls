VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Designer")
'@ModuleDescription("Interface of main class")
'@Interface

Option Explicit


'Get value from one range
Public Property Get ValueOf(ByVal tagName As String) As String: End Property
'Check if all the paths are ready for linelist generation
Public Property Get Ready() As Boolean: End Property

'Get the linelist password
Public Property Get LinelistPassWord() As String: End Property
'Add information to end user
Public Sub AddInfo(ByVal trads As IDesTranslation, ByVal msgCode As String, _
                   Optional ByVal infoName As String = "edition"): End Sub

'Exposed methods
Public Sub CheckReadiness(ByVal trads As IDesTranslation): End Sub
'Check if the linelist exists and ask for replacement
Public Sub CheckFileExistence(ByVal trads As IDesTranslation): End Sub
Public Sub UpdateStatus(ByVal numValue As Integer): End Sub
'Open a linelist
'Public Sub OpenLL(): End Sub
'Set all the input ranges to white
Public Sub ClearInputRanges(Optional ByVal clearValues As Boolean = False): End Sub
Public Sub OpenLL() : End Sub
'Check if the linelist could be built on a ribbon template or not
Public Property Get HasTemplate() As Boolean: End Property
Public Sub CheckRibbonExistence(ByVal trads As IDesTranslation): End Sub



