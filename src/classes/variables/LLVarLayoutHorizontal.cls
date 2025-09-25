VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarLayoutHorizontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Layout strategy for horizontal linelist variables")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarLayout

Private Type THorizontalLayout
    baseSheet As Worksheet
    printedSheet As Worksheet
    anchorIndex As Long
    valueCell As Range
    printedValueCell As Range
End Type

Private this As THorizontalLayout
Private Const CLASSNAME As String = "LLVarLayoutHorizontal"
Private Const DATA_START_ROW As Long = 8
Private Const LAYOUT_KEY As String = "horizontal"

'@section Lifecycle
'===============================================================================
Private Sub EnsureConfigured()
    If this.baseSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "Horizontal layout has not been configured with a worksheet"
    End If
    If this.anchorIndex <= 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "Horizontal layout anchor index is missing"
    End If
End Sub

Private Sub InitialiseValueCells()
    EnsureConfigured

    If this.valueCell Is Nothing Then
        Set this.valueCell = this.baseSheet.Cells(DATA_START_ROW + 1, this.anchorIndex)
    End If

    If (Not this.printedSheet Is Nothing) And (this.printedValueCell Is Nothing) Then
        Set this.printedValueCell = this.printedSheet.Cells(DATA_START_ROW + 1, this.anchorIndex)
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errorMessage As String)
    Err.Raise errNumber, CLASSNAME, errorMessage
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILLVarLayout_Configure(ByVal baseSheet As Worksheet, _
                                   ByVal anchorIndex As Long, _
                                   Optional ByVal printedSheet As Worksheet)

    If baseSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "Base worksheet is required for the horizontal layout"
    End If

    If anchorIndex <= 0 Then
        ThrowError ProjectError.InvalidArgument, _
            "Anchor column index must be greater than zero for the horizontal layout"
    End If

    Set this.baseSheet = baseSheet
    this.anchorIndex = anchorIndex
    Set this.printedSheet = printedSheet

    Set this.valueCell = Nothing
    Set this.printedValueCell = Nothing
End Sub

Private Sub ILLVarLayout_Reset()
    Set this.baseSheet = Nothing
    Set this.printedSheet = Nothing
    this.anchorIndex = 0
    Set this.valueCell = Nothing
    Set this.printedValueCell = Nothing
End Sub

Private Property Get ILLVarLayout_ValueCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_ValueCell = this.valueCell
End Property

Private Property Get ILLVarLayout_PrintedValueCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_PrintedValueCell = this.printedValueCell
End Property

Private Property Get ILLVarLayout_LabelCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_LabelCell = this.valueCell.Offset(-2, 0)
End Property

Private Property Get ILLVarLayout_SublabelTarget() As Range
    InitialiseValueCells
    Set ILLVarLayout_SublabelTarget = this.valueCell.Offset(-2, 0)
End Property

Private Property Get ILLVarLayout_NameCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_NameCell = this.valueCell.Offset(-1, 0)
End Property

Private Property Get ILLVarLayout_NoteTarget() As Range
    InitialiseValueCells
    Set ILLVarLayout_NoteTarget = this.valueCell.Offset(-2, 0)
End Property

Private Property Get ILLVarLayout_ControlCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_ControlCell = this.valueCell.Offset(-5, 0)
End Property

Private Property Get ILLVarLayout_AutoOriginCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_AutoOriginCell = this.valueCell.Offset(-6, 0)
End Property

Private Property Get ILLVarLayout_AnchorCell() As Range
    InitialiseValueCells
    Set ILLVarLayout_AnchorCell = this.valueCell.Offset(-2, 0)
End Property

Private Property Get ILLVarLayout_HiddenRange() As Range
    InitialiseValueCells
    Set ILLVarLayout_HiddenRange = this.valueCell.EntireColumn
End Property

Private Function ILLVarLayout_SupportsPrintedLayout() As Boolean
    InitialiseValueCells
    ILLVarLayout_SupportsPrintedLayout = Not (this.printedValueCell Is Nothing)
End Function

Private Property Get ILLVarLayout_LayoutKey() As String
    ILLVarLayout_LayoutKey = LAYOUT_KEY
End Property

