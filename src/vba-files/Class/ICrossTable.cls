VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICrossTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("Interface of crosstable")
'@IgnoreModule

Option Explicit

Public Property Get HeaderRange() As Range: End Property
Public Property Get RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, Optional ByVal onlyCategories As Boolean = False) As Range: End Property
Public Property Get ColumnRange(ByVal translationKey As String, Optional ByVal includeHeaders = False) As Range: End Property
Public Property Get RowRange(ByVal translationKey As String) As Range: End Property
Public Property Get StartRow() As Long: End Property
'Exposed methods

Public Sub AddRows(): End Sub
Public Sub AddColumns(): End Sub
Public Sub AddHeader(): End Sub
Public Sub NameRanges(): End Sub
Public Sub Build(): End Sub
