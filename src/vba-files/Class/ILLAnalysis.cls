VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILLAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Analysis Class"


'@Folder("Analysis")
'@ModuleDescription("Interface of Analysis Class")
'@IgnoreModule

Option Explicit

'Properties
Public Property Get Wksh() As Worksheet: End Property

'Exposed methods
Public Sub Import(ByVal Wksh As Worksheet): End Sub

Public Sub Translate(ByVal transObject As ITranslation): End Sub
