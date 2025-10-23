VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowHideExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Persist show/hide plans into workbook tables and reload them on demand"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("ShowHide")
'@ModuleDescription("Persist show/hide plans into workbook tables and reload them on demand")

Option Explicit
Implements IShowHideExport

Private Type TShowHideExport
    table As ICustomTable
End Type

Private this As TShowHideExport
Private Const CLASSNAME As String = "ShowHideExport"
Private Const COL_LAYER As String = "layer"
Private Const COL_FIELD_KEY As String = "field_key"
Private Const COL_HEADER_TEXT As String = "header_text"
Private Const COL_HIDDEN As String = "hidden_flag"


'@section Factory
'===============================================================================

Public Function Create(Optional ByVal table As ICustomTable) As IShowHideExport
    With New ShowHideExport
        If Not (table Is Nothing) Then
            .BindTable table
        End If
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

Public Property Get Self() As IShowHideExport
    Set Self = Me
End Property

Public Sub BindTable(ByVal table As ICustomTable)
    EnsureTableProvided table
    Set this.table = table
End Sub

Public Property Get BoundTable() As ICustomTable
    Set BoundTable = this.table
End Property


'@section Persistence
'===============================================================================

Public Sub ExportPlan(ByVal plan As IShowHidePlan)
    Dim targetLayer As ShowHideWorksheetLayer
    Dim lo As ListObject
    Dim actions As BetterArray
    Dim idx As Long
    Dim action As Variant
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim headerCol As Long
    Dim hiddenCol As Long
    Dim newRow As ListRow

    EnsureReady
    EnsurePlanProvided plan

    targetLayer = plan.TargetLayer
    ClearLayer targetLayer

    Set actions = plan.Actions
    If actions Is Nothing Then Exit Sub
    If actions.Length = 0 Then Exit Sub

    Set lo = BoundListObject()
    EnsureColumns lo
    layerCol = ColumnIndex(lo, COL_LAYER)
    fieldCol = ColumnIndex(lo, COL_FIELD_KEY)
    headerCol = ColumnIndex(lo, COL_HEADER_TEXT)
    hiddenCol = ColumnIndex(lo, COL_HIDDEN)

    For idx = actions.LowerBound To actions.UpperBound
        action = actions.Item(idx)
        Set newRow = lo.ListRows.Add
        With newRow.Range
            .Cells(1, layerCol).Value = LayerToken(targetLayer)
            .Cells(1, fieldCol).Value = action(0)
            .Cells(1, headerCol).Value = action(1)
            .Cells(1, hiddenCol).Value = BoolToString(CBool(action(2)))
        End With
    Next idx
End Sub

Public Function ImportPlan(ByVal layer As ShowHideWorksheetLayer) As IShowHidePlan
    Dim token As String
    Dim lo As ListObject
    Dim plan As IShowHidePlan
    Dim layerCol As Long
    Dim fieldCol As Long
    Dim headerCol As Long
    Dim hiddenCol As Long
    Dim currentRow As ListRow
    Dim currentLayer As String
    Dim fieldKey As String
    Dim headerText As String
    Dim hiddenFlag As Boolean

    EnsureReady

    token = LayerToken(layer)
    Set plan = ShowHidePlan.Create(layer)
    Set lo = BoundListObject()

    layerCol = ColumnIndex(lo, COL_LAYER)
    fieldCol = ColumnIndex(lo, COL_FIELD_KEY)
    headerCol = ColumnIndex(lo, COL_HEADER_TEXT)
    hiddenCol = ColumnIndex(lo, COL_HIDDEN)

    If lo.ListRows.Count = 0 Then
        Set ImportPlan = plan
        Exit Function
    End If

    For Each currentRow In lo.ListRows
        currentLayer = CStr(currentRow.Range.Cells(1, layerCol).Value)
        If StrComp(currentLayer, token, vbTextCompare) = 0 Then
            fieldKey = Trim$(CStr(currentRow.Range.Cells(1, fieldCol).Value))
            headerText = Trim$(CStr(currentRow.Range.Cells(1, headerCol).Value))
            hiddenFlag = ToBoolean(currentRow.Range.Cells(1, hiddenCol).Value)
            If LenB(fieldKey) > 0 Then
                plan.AddVisibility fieldKey, headerText, hiddenFlag
            End If
        End If
    Next currentRow

    Set ImportPlan = plan
End Function

Public Sub ClearLayer(ByVal layer As ShowHideWorksheetLayer)
    Dim lo As ListObject
    Dim token As String
    Dim layerCol As Long
    Dim idx As Long

    EnsureReady

    token = LayerToken(layer)
    Set lo = BoundListObject()

    layerCol = ColumnIndex(lo, COL_LAYER)
    If lo.ListRows.Count = 0 Then Exit Sub

    For idx = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(lo.ListRows(idx).Range.Cells(1, layerCol).Value), token, vbTextCompare) = 0 Then
            lo.ListRows(idx).Delete
        End If
    Next idx
End Sub


'@section Helpers
'===============================================================================

Private Sub EnsureReady()
    If this.table Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "ShowHideExport requires a bound table before persisting data"
    End If
End Sub

Private Sub EnsureTableProvided(ByVal table As ICustomTable)
    If table Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Cannot bind an empty table reference"
    End If
    If TypeName(table) <> "CustomTable" Then
        ThrowError ProjectError.InvalidArgument, _
                  "ShowHideExport expects a CustomTable implementation"
    End If
End Sub

Private Sub EnsurePlanProvided(ByVal plan As IShowHidePlan)
    If plan Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Cannot export an empty plan"
    End If
End Sub

Private Function BoundListObject() As ListObject
    Dim concrete As CustomTable

    Set concrete = this.table
    Set BoundListObject = concrete.Table
End Function

Private Function ColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim headers As Range
    Dim found As Range

    If lo Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Cannot resolve header index on an empty table"
    End If

    Set headers = lo.HeaderRowRange
    If headers Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                  "Header row is missing on the bound table"
    End If

    On Error Resume Next
        Set found = headers.Find(What:=headerName, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If found Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                  "Column '" & headerName & "' is missing from the bound table"
    End If

    ColumnIndex = found.Column - headers.Column + 1
End Function

Private Sub EnsureColumns(ByVal lo As ListObject)
    Dim required As Variant
    Dim idx As Long
    Dim name As String
    Dim headers As Range
    Dim existing As Range

    required = Array(COL_LAYER, COL_FIELD_KEY, COL_HEADER_TEXT, COL_HIDDEN)
    Set headers = lo.HeaderRowRange

    For idx = LBound(required) To UBound(required)
        name = CStr(required(idx))
        On Error Resume Next
            Set existing = headers.Find(What:=name, LookAt:=xlWhole, MatchCase:=False)
        On Error GoTo 0
        If existing Is Nothing Then
            ThrowError ProjectError.InvalidArgument, _
                      "The bound table '" & lo.Name & "' does not expose expected column '" & name & "'"
        End If
        Set existing = Nothing
    Next idx
End Sub

Private Function LayerToken(ByVal layer As ShowHideWorksheetLayer) As String
    Select Case layer
        Case ShowHideLayerHList
            LayerToken = "hlist"
        Case ShowHideLayerCRF
            LayerToken = "crf"
        Case ShowHideLayerPrinted
            LayerToken = "printed"
        Case ShowHideLayerVList
            LayerToken = "vlist"
        Case ShowHideLayerDictionary
            LayerToken = "dictionary"
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                      "Unsupported show/hide layer value: " & CStr(layer)
    End Select
End Function

Private Function BoolToString(ByVal value As Boolean) As String
    If value Then
        BoolToString = "true"
    Else
        BoolToString = "false"
    End If
End Function

Private Function ToBoolean(ByVal value As Variant) As Boolean
    If VarType(value) = vbBoolean Then
        ToBoolean = CBool(value)
    ElseIf VarType(value) = vbString Then
        Select Case LCase$(Trim$(CStr(value)))
            Case "true", "yes", "1"
                ToBoolean = True
            Case Else
                ToBoolean = False
        End Select
    ElseIf IsNumeric(value) Then
        ToBoolean = (CLng(value) <> 0)
    Else
        ToBoolean = False
    End If
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface
'===============================================================================

Private Sub IShowHideExport_BindTable(ByVal table As ICustomTable)
    BindTable table
End Sub

Private Property Get IShowHideExport_BoundTable() As ICustomTable
    Set IShowHideExport_BoundTable = BoundTable
End Property

Private Sub IShowHideExport_ExportPlan(ByVal plan As IShowHidePlan)
    ExportPlan plan
End Sub

Private Function IShowHideExport_ImportPlan(ByVal layer As ShowHideWorksheetLayer) As IShowHidePlan
    Set IShowHideExport_ImportPlan = ImportPlan(layer)
End Function

Private Sub IShowHideExport_ClearLayer(ByVal layer As ShowHideWorksheetLayer)
    ClearLayer layer
End Sub
