VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISetupImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of setup import"

'@Folder("setup file")
'@ModuleDescription("Interface of setup import")
'@Interface

Option Explicit

'Exposed methods

Public Sub check(ByVal impDict As Boolean, _
  ByVal impChoi As Boolean, _
  ByVal impExp As Boolean, _
  ByVal impAna As Boolean, _
  ByVal impTrans As Boolean, _
  Optional ByVal cleanSetup As Boolean = False): End Sub

Public Sub Import(ByVal pass As IPasswords, ByVal sheetsList As BetterArray): End Sub
Public Sub Clean(ByVal pass As IPasswords, ByVal sheetsList As BetterArray): End Sub
