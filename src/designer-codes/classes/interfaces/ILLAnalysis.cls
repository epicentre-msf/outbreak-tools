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

'@Interface
'@Folder("Analysis")
'@ModuleDescription("Interface of Analysis Class")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Properties
Public Property Get Wksh() As Worksheet: End Property

'Exposed methods
Public Sub Import(ByVal Wksh As Worksheet): End Sub

Public Sub Translate(ByVal TransObject As ITranslationObject): End Sub
Public Sub Build(ll As ILinelist): End Sub
