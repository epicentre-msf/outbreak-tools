VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IOutputSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Linelist")
'@ModuleDescription("Interface of Output yjr ")

Option Explicit

'Exposed methods

Public Sub Prepare() : End Sub
'Add Dictionary, Translations or Analysis
Public Sub AddData(): End Sub
Public Sub Save(ByVal trads As ITranslation, Optional ByVal useFilters As Boolean = False): End Sub
Public Property Get ExportFolder() As String : End Property
Public Sub SaveGeo(ByVal geoObj As ILLGeo, Optional ByVal onlyHistoric As Boolean = False): End Sub
Public Sub CloseAll(): End Sub
