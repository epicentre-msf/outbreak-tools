VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IImpSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Import Specs"

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@Folder("Linelist")
'@Interface
'@ModuleDescription("Interface of Import Specs")

Option Explicit

'Exposed methods

'Import a geobase
Public Sub ImportGeobase(Optional ByVal histoOnly As Boolean = False)
End Sub

'Import data for migrations
Public Sub ImportMigration(): End Sub

'Clear data in the current linelist
Public Sub ClearData(): End Sub

'Show the import report
Public Sub ShowReport(): End Sub
