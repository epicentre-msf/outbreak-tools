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
'@IgnoreModule

Option Explicit


Public Property Get OutputPath() As String: End Property 'Linelist Directory
Public Property Get LinelistName() As String: End Property 'Linelist Name
Public Property Get SetupPath() As String: End Property 'Setup Path
Public Property Get GeoPath() As String: End Property ' GeoPath
'Check if all the paths are ready for linelist generation
Public Property Get Ready() As Boolean: End Property

'Exposed methods
Public Sub CheckReadiness(ByVal Trads As IDesTranslation): End Sub
Public Sub UpdateStatus(ByVal numValue As Integer): End Sub
