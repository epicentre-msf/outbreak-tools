VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarLayoutCRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Layout strategy for CRF variables")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarLayout

Private Type TCRFLayout
    baseSheet As Worksheet
    anchorRow As Long
    valueCell As Range
End Type

Private this As TCRFLayout
Private Const CLASSNAME As String = "LLVarLayoutCRF"
Private Const LAYOUT_KEY As String = "crf"
Private Const DATA_COLUMN As Long = 1

'@section Guards
'===============================================================================
Private Sub EnsureConfigured()
    If this.baseSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "CRF layout must be configured with a worksheet before use"
    End If
    If this.anchorRow <= 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "CRF layout anchor row is missing"
    End If
End Sub

Private Sub EnsureValueCell()
    EnsureConfigured
    If this.valueCell Is Nothing Then
        Set this.valueCell = this.baseSheet.Cells(this.anchorRow, DATA_COLUMN)
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Sub ILLVarLayout_Configure(ByVal baseSheet As Worksheet, _
                                   ByVal anchorIndex As Long, _
                                   Optional ByVal printedSheet As Worksheet)

    If baseSheet Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Base worksheet is required for the CRF layout"
    End If

    If anchorIndex <= 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "CRF layout requires a positive anchor row"
    End If

    Set this.baseSheet = baseSheet
    this.anchorRow = anchorIndex
    Set this.valueCell = Nothing
End Sub

Private Sub ILLVarLayout_Reset()
    Set this.baseSheet = Nothing
    this.anchorRow = 0
    Set this.valueCell = Nothing
End Sub

Private Property Get ILLVarLayout_ValueCell() As Range
    EnsureValueCell
    Set ILLVarLayout_ValueCell = this.valueCell
End Property

Private Property Get ILLVarLayout_PrintedValueCell() As Range
    Set ILLVarLayout_PrintedValueCell = Nothing
End Property

Private Property Get ILLVarLayout_LabelCell() As Range
    EnsureValueCell
    Set ILLVarLayout_LabelCell = this.valueCell
End Property

Private Property Get ILLVarLayout_SublabelTarget() As Range
    EnsureValueCell
    Set ILLVarLayout_SublabelTarget = this.valueCell
End Property

Private Property Get ILLVarLayout_NameCell() As Range
    Set ILLVarLayout_NameCell = Nothing
End Property

Private Property Get ILLVarLayout_NoteTarget() As Range
    Set ILLVarLayout_NoteTarget = Nothing
End Property

Private Property Get ILLVarLayout_ControlCell() As Range
    Set ILLVarLayout_ControlCell = Nothing
End Property

Private Property Get ILLVarLayout_AutoOriginCell() As Range
    Set ILLVarLayout_AutoOriginCell = Nothing
End Property

Private Property Get ILLVarLayout_AnchorCell() As Range
    EnsureValueCell
    Set ILLVarLayout_AnchorCell = this.valueCell
End Property

Private Property Get ILLVarLayout_HiddenRange() As Range
    EnsureValueCell
    Set ILLVarLayout_HiddenRange = this.valueCell.EntireRow
End Property

Private Function ILLVarLayout_SupportsPrintedLayout() As Boolean
    ILLVarLayout_SupportsPrintedLayout = False
End Function

Private Property Get ILLVarLayout_LayoutKey() As String
    ILLVarLayout_LayoutKey = LAYOUT_KEY
End Property
