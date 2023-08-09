VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChoiceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of choice formula"


'@Interface
'@Folder("Custom Functions")
'@ModuleDescription("Interface of choice formula")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Public Property Get parsedFormula() As String
End Property

'List of categories
Public Property Get Categories() As BetterArray
End Property

'Validity test
Public Property Get Valid() As Boolean
End Property

'Get the choice name
'@Ignored ProcedureNotUsed
Public Property Get choiceName() As String
End Property
