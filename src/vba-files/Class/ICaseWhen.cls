VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICaseWhen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Dictionary")
'@ModuleDescription("Interface of CaseWhen")
'@IgnoreModule

Option Explicit

Public Property Get Valid() As Boolean: End Property
Public Property Get parsedFormula() As String: End Property

Public Property Get Categories() As BetterArray: End Property
