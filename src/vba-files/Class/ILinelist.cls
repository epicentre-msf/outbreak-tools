VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILinelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Interface of Linelist")
'@IgnoreModule

Option Explicit

Public Property Get Wksh(ByVal sheetName As String) As Worksheet: End Property
Public Property Get Dropdown() As IDropdownLists: End Property
Public Property Get linelistData() As ILinelistSpecs: End Property
Public Property Get Dictionary() As ILLdictionary: End Property


'Exposed methods
Public Sub Prepare(): End Sub
Public Sub Build(): End Sub
Public Function sheetExists(ByVal sheetName As String) As Boolean: End Function