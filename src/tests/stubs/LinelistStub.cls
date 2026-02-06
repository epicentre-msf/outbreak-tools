VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Minimal ILinelist double for ListBuilder factory tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILinelist

Private specs As ILinelistSpecs
Private dict As ILLdictionary

'@section Configuration
'===============================================================================

Public Sub Initialise(ByVal specsObj As ILinelistSpecs, ByVal dictObj As ILLdictionary)
    Set specs = specsObj
    Set dict = dictObj
End Sub

'@section ILinelist implementation
'===============================================================================

Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Nothing
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = Nothing
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = Nothing
End Property

Private Function ILinelist_SheetExists(ByVal sheetName As String) As Boolean
    ILinelist_SheetExists = False
End Function

Private Sub ILinelist_AddOutputSheet(ByVal sheetName As String, _
                                     Optional ByVal Hide As Long = xlSheetVeryHidden, _
                                     Optional ByVal sheetScope As Byte = 1)
End Sub

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Set ILinelist_Dropdown = Nothing
End Property

Private Property Get ILinelist_Pivots() As ICustomPivotTables
    Set ILinelist_Pivots = Nothing
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = specs
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = dict
End Property

Private Sub ILinelist_Prepare()
End Sub

Private Sub ILinelist_SaveLL()
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub
