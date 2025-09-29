VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarLayoutVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Layout strategy for vertical linelist variables")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarLayout

Private Type TVerticalLayout
    baseSheet As Worksheet
    anchorIndex As Long
    valueCell As Range
End Type

Private this As TVerticalLayout
Private Const CLASSNAME As String = "LLVarLayoutVertical"
Private Const DATA_START_COLUMN As Long = 5
Private Const LAYOUT_KEY As String = "vertical"

'@section Lifecycle
'===============================================================================
Private Sub EnsureConfigured()
    If this.baseSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "Vertical layout has not been configured with a worksheet"
    End If
    If this.anchorIndex <= 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
            "Vertical layout anchor index is missing"
    End If
End Sub

Private Sub InitialiseValueCell()
    EnsureConfigured
    If this.valueCell Is Nothing Then
        Set this.valueCell = this.baseSheet.Cells(this.anchorIndex, DATA_START_COLUMN)
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
            "Base worksheet is required for the vertical layout"
    End If

    If anchorIndex <= 0 Then
        ThrowError ProjectError.InvalidArgument, _
            "Anchor row index must be greater than zero for the vertical layout"
    End If

    Set this.baseSheet = baseSheet
    this.anchorIndex = anchorIndex
    Set this.valueCell = Nothing
End Sub

Private Sub ILLVarLayout_Reset()
    Set this.baseSheet = Nothing
    this.anchorIndex = 0
    Set this.valueCell = Nothing
End Sub

Private Property Get ILLVarLayout_ValueCell() As Range
    InitialiseValueCell
    Set ILLVarLayout_ValueCell = this.valueCell
End Property

Private Property Get ILLVarLayout_PrintedValueCell() As Range
    Set ILLVarLayout_PrintedValueCell = Nothing
End Property

Private Property Get ILLVarLayout_LabelCell() As Range
    InitialiseValueCell
    Set ILLVarLayout_LabelCell = this.valueCell.Offset(0, -1)
End Property

Private Property Get ILLVarLayout_SublabelTarget() As Range
    InitialiseValueCell
    Set ILLVarLayout_SublabelTarget = this.valueCell.Offset(0, -1)
End Property

Private Property Get ILLVarLayout_NameCell() As Range
    InitialiseValueCell
    Set ILLVarLayout_NameCell = this.valueCell
End Property

Private Property Get ILLVarLayout_NoteTarget() As Range
    InitialiseValueCell
    Set ILLVarLayout_NoteTarget = this.valueCell
End Property

Private Property Get ILLVarLayout_ControlCell() As Range
    InitialiseValueCell
    Set ILLVarLayout_ControlCell = this.valueCell.Offset(0, 1)
End Property

Private Property Get ILLVarLayout_AutoOriginCell() As Range
    Set ILLVarLayout_AutoOriginCell = Nothing
End Property

Private Property Get ILLVarLayout_AnchorCell() As Range
    InitialiseValueCell
    Set ILLVarLayout_AnchorCell = this.valueCell
End Property

Private Property Get ILLVarLayout_HiddenRange() As Range
    InitialiseValueCell
    Set ILLVarLayout_HiddenRange = this.valueCell.EntireRow
End Property

Private Function ILLVarLayout_SupportsPrintedLayout() As Boolean
    ILLVarLayout_SupportsPrintedLayout = False
End Function

Private Property Get ILLVarLayout_LayoutKey() As String
    ILLVarLayout_LayoutKey = LAYOUT_KEY
End Property

