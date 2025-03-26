VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of translation"

'@Folder("Dictionary")
'@ModuleDescription("Interface of translation")
'@Interface

Option Explicit

'Exposed methods
Public Sub UpdateTrans(ByVal rngSh As Worksheet): End Sub
Public Sub AddTransLang(ByVal langVal As String): End Sub
