VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVListSubSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Interface
'@Folder("Dictionary")
'@ModuleDescription("Interface of ")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

'Exposed methods
Public Property Get HasNext() As Boolean: End Property
Public Property Get NextSubSection() As IVListSubSections: End Property

'Build the sub sections
Public Sub Build():  End Sub
  
'Format the subSection
Public Sub Format(): End Sub
