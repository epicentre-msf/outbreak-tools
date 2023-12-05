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
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit

Public Property Get Wksh(ByVal sheetName As String) As Worksheet
End Property

Public Property Get PrintWksh(ByVal sheetName As String) As Worksheet
End Property

Public Property Get Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
End Property

Public Property Get Pivots() As ICustomPivotTables
End Property

Public Property Get LinelistData() As ILinelistSpecs: End Property
Public Property Get Dictionary() As ILLdictionary: End Property

'Exposed methods
Public Sub Prepare(): End Sub

Public Function sheetExists(ByVal sheetName As String) As Boolean: End Function
'Add output sheet
Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal usePrinted As Boolean = False)
End Sub

Public Sub SaveLL(): End Sub

'Error management
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString): End Sub
