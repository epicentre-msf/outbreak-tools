VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphTablesSpecsStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub implementation of ITablesSpecs for graph builder unit tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation, SuperfluousAnnotationArgument

Option Explicit

Implements ITablesSpecs

Private Type TSpecState
    tableTypeCode As Byte
    tableId As String
    sectionId As String
    hasPercentage As Boolean
    hasMissing As Boolean
    hasTotal As Boolean
    hasGraph As Boolean
    isNewSection As Boolean
    previousSpecs As ITablesSpecs
    rowCategories As BetterArray
    columnCategories As BetterArray
    keys As BetterArray
    values As BetterArray
End Type

Private this As TSpecState

Private Sub Class_Initialize()
    this.hasGraph = True
End Sub

Public Sub Configure(ByVal tableTypeCode As Byte, _
                     ByVal tableId As String, _
                     Optional ByVal sectionId As String = vbNullString)
    this.tableTypeCode = tableTypeCode
    this.tableId = tableId
    this.sectionId = sectionId
End Sub

Public Sub SetHasPercentage(ByVal hasPercentage As Boolean)
    this.hasPercentage = hasPercentage
End Sub

Public Sub SetHasMissing(ByVal hasMissing As Boolean)
    this.hasMissing = hasMissing
End Sub

Public Sub SetHasTotal(ByVal hasTotal As Boolean)
    this.hasTotal = hasTotal
End Sub

Public Sub SetHasGraph(ByVal hasGraph As Boolean)
    this.hasGraph = hasGraph
End Sub

Public Sub SetIsNewSection(ByVal isNew As Boolean)
    this.isNewSection = isNew
End Sub

Public Sub SetPrevious(ByVal previousSpecs As ITablesSpecs)
    Set this.previousSpecs = previousSpecs
End Sub

Public Sub SetRowCategories(ParamArray categories() As Variant)
    Dim idx As Long

    Set this.rowCategories = New BetterArray
    this.rowCategories.LowerBound = 1

    For idx = LBound(categories) To UBound(categories)
        this.rowCategories.Push categories(idx)
    Next idx
End Sub

Public Sub SetColumnCategories(ParamArray categories() As Variant)
    Dim idx As Long

    Set this.columnCategories = New BetterArray
    this.columnCategories.LowerBound = 1

    For idx = LBound(categories) To UBound(categories)
        this.columnCategories.Push categories(idx)
    Next idx
End Sub

Public Sub SetValue(ByVal columnName As String, ByVal valueText As String)
    Dim keyText As String
    Dim idx As Long

    keyText = LCase$(Trim$(columnName))

    If this.keys Is Nothing Then
        Set this.keys = New BetterArray
        Set this.values = New BetterArray
        this.keys.LowerBound = 1
        this.values.LowerBound = 1
    End If

    If this.keys.Length > 0 Then
        For idx = this.keys.LowerBound To this.keys.UpperBound
            If LCase$(CStr(this.keys.Item(idx))) = keyText Then Exit Sub
        Next idx
    End If

    this.keys.Push keyText
    this.values.Push valueText
End Sub

Public Property Get Self() As ITablesSpecs
    Set Self = Me
End Property

'Interface implementation -------------------------------------------------------
Private Property Get ITablesSpecs_TableType() As AnalysisTablesType
    ITablesSpecs_TableType = this.tableTypeCode
End Property

Private Property Get ITablesSpecs_Dictionary() As ILLdictionary
    Set ITablesSpecs_Dictionary = Nothing
End Property

Private Property Get ITablesSpecs_RowCategories() As BetterArray
    Set ITablesSpecs_RowCategories = this.rowCategories
End Property

Private Property Get ITablesSpecs_ColumnCategories() As BetterArray
    Set ITablesSpecs_ColumnCategories = this.columnCategories
End Property

Private Property Get ITablesSpecs_isNewSection() As Boolean
    ITablesSpecs_isNewSection = this.isNewSection
End Property

Private Property Get ITablesSpecs_TableId() As String
    ITablesSpecs_TableId = this.tableId
End Property

Private Property Get ITablesSpecs_ValidTable() As Boolean
    ITablesSpecs_ValidTable = True
End Property

Private Property Get ITablesSpecs_HasMissing() As Boolean
    ITablesSpecs_HasMissing = this.hasMissing
End Property

Private Property Get ITablesSpecs_HasTotal() As Boolean
    ITablesSpecs_HasTotal = this.hasTotal
End Property

Private Property Get ITablesSpecs_HasGraph() As Boolean
    ITablesSpecs_HasGraph = this.hasGraph
End Property

Private Property Get ITablesSpecs_HasPercentage() As Boolean
    ITablesSpecs_HasPercentage = this.hasPercentage
End Property

Private Property Get ITablesSpecs_SpatialTableTypes() As String
    ITablesSpecs_SpatialTableTypes = vbNullString
End Property

Private Property Get ITablesSpecs_Previous() As ITablesSpecs
    Set ITablesSpecs_Previous = this.previousSpecs
End Property

Private Property Get ITablesSpecs_NextSpecs(ByVal anchorRng As Range) As ITablesSpecs
    Set ITablesSpecs_NextSpecs = Nothing
End Property

Private Property Get ITablesSpecs_TableSectionId() As String
    ITablesSpecs_TableSectionId = this.sectionId
End Property

Private Function ITablesSpecs_Value(ByVal colName As String) As String
    Dim keyText As String
    Dim idx As Long

    keyText = LCase$(Trim$(colName))

    If this.keys Is Nothing Then
        ITablesSpecs_Value = vbNullString
        Exit Function
    End If

    If this.keys.Length > 0 Then
        For idx = this.keys.LowerBound To this.keys.UpperBound
            If LCase$(CStr(this.keys.Item(idx))) = keyText Then
                ITablesSpecs_Value = CStr(this.values.Item(idx))
                Exit Function
            End If
        Next idx
    End If

    ITablesSpecs_Value = vbNullString
End Function
