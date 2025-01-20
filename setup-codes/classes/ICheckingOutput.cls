VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICheckingOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of checking output"


'@Folder("Checks")
'@ModuleDescription("Interface of checking output")
'@Interface

Option Explicit

'Exposed methods

Public Sub PrintOutput(ByVal checkTable As BetterArray): End Sub
