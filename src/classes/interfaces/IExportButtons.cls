VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "IExportButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Dictionary")
'@ModuleDescription("Interface of Export button")
'@IgnoreModule

Option Explicit

'Exposed methods

Public Event ClickExport(ByVal exportNumber As Long)


