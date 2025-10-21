VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Minimal ICrossTable test stub for graph engine scenarios")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ICrossTable

Private Type TCrossTableStubState
    specification As ITablesSpecs
End Type

Private this As TCrossTableStubState

Public Sub Configure(ByVal specification As ITablesSpecs)
    Set this.specification = specification
End Sub

Private Property Get ICrossTable_HeaderRange() As Range
    Set ICrossTable_HeaderRange = Nothing
End Property

Private Property Get ICrossTable_RowsCategoriesRange(Optional ByVal includeHeaders As Boolean = True, Optional ByVal onlyCategories As Boolean = False) As Range
    Set ICrossTable_RowsCategoriesRange = Nothing
End Property

Private Property Get ICrossTable_ColumnRange(ByVal colVal As String, Optional ByVal onlyCategories As Boolean = False, Optional ByVal includeHeaders As Boolean = False) As Range
    Set ICrossTable_ColumnRange = Nothing
End Property

Private Property Get ICrossTable_RowRange(ByVal rowVal As String, Optional ByVal includeHeaders As Boolean = True) As Range
    Set ICrossTable_RowRange = Nothing
End Property

Private Property Get ICrossTable_StartRow() As Long
    ICrossTable_StartRow = 0
End Property

Private Property Get ICrossTable_Previous() As ICrossTable
    Set ICrossTable_Previous = Nothing
End Property

Private Property Get ICrossTable_Specifications() As ITablesSpecs
    Set ICrossTable_Specifications = this.specification
End Property

Private Property Get ICrossTable_Wksh() As Worksheet
    Set ICrossTable_Wksh = Nothing
End Property

Private Property Get ICrossTable_NumberOfColumns() As Long
    ICrossTable_NumberOfColumns = 0
End Property

Private Property Get ICrossTable_EndColumn() As Long
    ICrossTable_EndColumn = 0
End Property

Private Property Get ICrossTable_EndRow() As Long
    ICrossTable_EndRow = 0
End Property

Private Property Get ICrossTable_TimeSeriesHeader() As String
    ICrossTable_TimeSeriesHeader = vbNullString
End Property

Private Property Get ICrossTable_NamedRangesList() As BetterArray
    Set ICrossTable_NamedRangesList = Nothing
End Property

Private Sub ICrossTable_AddRows()
End Sub

Private Sub ICrossTable_AddColumns()
End Sub

Private Sub ICrossTable_AddHeader()
End Sub

Private Sub ICrossTable_NameRanges()
End Sub

Private Sub ICrossTable_Build()
End Sub

Private Sub ICrossTable_Format(ByVal DesignFormat As ILLFormat)
End Sub
