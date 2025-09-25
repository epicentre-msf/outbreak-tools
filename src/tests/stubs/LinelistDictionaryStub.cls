VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistDictionaryStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Minimal ILinelist implementation exposing a dictionary reference")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ILinelist

Private Type TLinelistState
    dictionary As ILLdictionary
End Type

Private this As TLinelistState

Public Sub Initialise(ByVal dictionary As ILLdictionary)
    Set this.dictionary = dictionary
End Sub

Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Nothing
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = Nothing
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = Nothing
End Property

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Set ILinelist_Dropdown = Nothing
End Property

Private Property Get ILinelist_Pivots() As ICustomPivotTables
    Set ILinelist_Pivots = Nothing
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = Nothing
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = this.dictionary
End Property

Private Sub ILinelist_Prepare()
End Sub

Private Function ILinelist_sheetExists(ByVal sheetName As String) As Boolean
    ILinelist_sheetExists = False
End Function

Private Sub ILinelist_AddOutputSheet(ByVal sheetName As String, Optional ByVal Hide As Long = xlSheetVeryHidden, Optional ByVal sheetScope As Byte = 1)
End Sub

Private Sub ILinelist_SaveLL()
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
End Sub

