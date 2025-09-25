VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSpecsNavigatorSpecStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of ITablesSpecs for navigator tests")
'@IgnoreModule UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITablesSpecs

Private Type TSpecStub
    validFlag As Boolean
    tableType As AnalysisTablesType
    tableRange As Range
End Type

Private this As TSpecStub

Public Sub Configure(ByVal isValid As Boolean, _
                     ByVal tableRange As Range, _
                     Optional ByVal tableType As AnalysisTablesType = TypeUnivariate)
    this.validFlag = isValid
    this.tableType = tableType
    Set this.tableRange = tableRange
End Sub

Public Property Get Self() As ITablesSpecs
    Set Self = Me
End Property

Public Property Get RowNumber() As Long
    If this.tableRange Is Nothing Then
        RowNumber = 0
    Else
        RowNumber = this.tableRange.Row
    End If
End Property


'Interface implementation -------------------------------------------------------
Private Property Get ITablesSpecs_TableType() As AnalysisTablesType
    ITablesSpecs_TableType = this.tableType
End Property

Private Property Get ITablesSpecs_Dictionary() As ILLdictionary
    Set ITablesSpecs_Dictionary = Nothing
End Property

Private Property Get ITablesSpecs_RowCategories() As BetterArray
    Set ITablesSpecs_RowCategories = Nothing
End Property

Private Property Get ITablesSpecs_ColumnCategories() As BetterArray
    Set ITablesSpecs_ColumnCategories = Nothing
End Property

Private Property Get ITablesSpecs_isNewSection() As Boolean
    ITablesSpecs_isNewSection = False
End Property

Private Property Get ITablesSpecs_TableId() As String
    ITablesSpecs_TableId = vbNullString
End Property

Private Property Get ITablesSpecs_ValidTable() As Boolean
    ITablesSpecs_ValidTable = this.validFlag
End Property

Private Property Get ITablesSpecs_HasMissing() As Boolean
    ITablesSpecs_HasMissing = False
End Property

Private Property Get ITablesSpecs_HasTotal() As Boolean
    ITablesSpecs_HasTotal = False
End Property

Private Property Get ITablesSpecs_HasGraph() As Boolean
    ITablesSpecs_HasGraph = False
End Property

Private Property Get ITablesSpecs_HasPercentage() As Boolean
    ITablesSpecs_HasPercentage = False
End Property

Private Property Get ITablesSpecs_SpatialTableTypes() As String
    ITablesSpecs_SpatialTableTypes = vbNullString
End Property

Private Property Get ITablesSpecs_Previous() As ITablesSpecs
    Set ITablesSpecs_Previous = Nothing
End Property

Private Property Get ITablesSpecs_NextSpecs(ByVal anchorRng As Range) As ITablesSpecs
    Set ITablesSpecs_NextSpecs = Nothing
End Property

Private Property Get ITablesSpecs_TableSectionId() As String
    ITablesSpecs_TableSectionId = vbNullString
End Property

Private Function ITablesSpecs_Value(ByVal colName As String) As String
    ITablesSpecs_Value = vbNullString
End Function

