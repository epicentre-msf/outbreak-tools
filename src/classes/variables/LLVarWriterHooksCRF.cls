VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterHooksCRF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("CRF-specific hooks for the reusable linelist writer")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriterHooks

Private Const CLASSNAME As String = "LLVarWriterHooksCRF"
Private Const HIDDEN_SCOPE As Byte = LinelistHiddenCell
Private Const MAIN_LABEL_SCOPE As Byte = HListMainLab
Private Const SUB_LABEL_SCOPE As Byte = HListSublab

'@section Helpers
'===============================================================================
Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

Private Function ResolveCRFRow(ByVal context As ILLVarContext) As Long
    Dim crfIndexText As String
    Dim crfIndexValue As Long

    crfIndexText = Trim$(context.ValueOf("crf index"))

    On Error Resume Next
        crfIndexValue = CLng(crfIndexText)
    On Error GoTo 0

    If crfIndexValue <= 0 Then
        ThrowError ProjectError.InvalidArgument, _
                   "Dictionary entry 'crf index' must resolve to a positive number"
    End If

    ResolveCRFRow = crfIndexValue
End Function

Private Function EnsureCRFWorksheet(ByVal context As ILLVarContext) As Worksheet
    Dim ll As ILinelist
    Dim sheetName As String
    Dim tableName As String
    Dim design As ILLFormat
    Dim targetSheet As Worksheet

    Set ll = context.Linelist
    sheetName = context.ValueOf("sheet name")
    tableName = context.ValueOf("table name")

    If LenB(sheetName) = 0 Then
        ThrowError ProjectError.ElementNotFound, _
                   "CRF worksheet cannot be resolved because the sheet name is missing"
    End If

    On Error Resume Next
        Set targetSheet = ll.CRFWksh(sheetName)
    On Error GoTo 0

    If targetSheet Is Nothing Then
        ll.AddOutputSheet sheetName, xlSheetHidden, sheetScope:=3

        On Error Resume Next
            Set targetSheet = ll.CRFWksh(sheetName)
        On Error GoTo 0

        If targetSheet Is Nothing Then
            ThrowError ProjectError.ElementNotFound, _
                       "Failed to create CRF worksheet '" & sheetName & "'"
        End If

        Set design = context.Design
        If Not design Is Nothing Then
            design.ApplyFormat targetSheet, scope:=HListCRFSheet
            design.ApplyFormat targetSheet.Cells(1, 3), scope:=HIDDEN_SCOPE
            design.ApplyFormat targetSheet.Cells(1, 4), scope:=HIDDEN_SCOPE
        End If

        targetSheet.Cells(1, 3).Value = "HList CRF"
        targetSheet.Cells(1, 4).Value = "crf_" & tableName
    End If

    Set EnsureCRFWorksheet = targetSheet
End Function

Private Function ShouldHideRow(ByVal statusValue As String, ByVal controlValue As String) As Boolean
    Dim loweredStatus As String
    Dim loweredControl As String

    loweredStatus = LCase$(Trim$(statusValue))
    loweredControl = LCase$(Trim$(controlValue))

    ShouldHideRow = (loweredStatus = "hidden") Or _
                    (loweredStatus = "optional, hidden") Or _
                    (loweredControl = "formula") Or _
                    (loweredControl = "choice_formula") Or _
                    (loweredControl = "case_when")
End Function

Private Sub ApplyLabelFormatting(ByVal context As ILLVarContext, _
                                 ByVal labelCell As Range, _
                                 ByVal mainLabelLength As Long)
    Dim design As ILLFormat

    If labelCell Is Nothing Then Exit Sub

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat labelCell.Characters(Start:=1, Length:=mainLabelLength), _
                           scope:=MAIN_LABEL_SCOPE
        If mainLabelLength < Len(labelCell.Value) Then
            design.ApplyFormat labelCell.Characters(Start:=mainLabelLength + 1), _
                               scope:=SUB_LABEL_SCOPE
        End If
    On Error GoTo 0
End Sub

Private Sub ApplyCRFTypeFormat(ByVal context As ILLVarContext, _
                               ByVal valueCell As Range)
    Dim varType As String
    Dim varFormat As String
    Dim controlType As String
    Dim scope As Byte
    Dim decimalsCount As Long
    Dim design As ILLFormat

    If valueCell Is Nothing Then Exit Sub

    varType = LCase$(Trim$(context.ValueOf("variable type")))
    varFormat = LCase$(Trim$(context.ValueOf("variable format")))
    controlType = LCase$(Trim$(context.ValueOf("control")))

    If controlType = "choice_manual" Then Exit Sub

    scope = 0
    decimalsCount = 0

    Select Case varType
    Case "text"
        scope = CRFText

    Case "date"
        If LenB(varFormat) = 0 Then
            scope = CRFDate
        Else
            scope = CRFFormatCustom
        End If

    Case "decimal", "integer"
        If LenB(varFormat) = 0 Then
            scope = IIf(varType = "decimal", CRFDecimal, CRFInteger)
        ElseIf varFormat = "integer" Or varFormat = "round0" Then
            scope = CRFInteger
        ElseIf varFormat = "percentage0" Then
            scope = CRFPercentage
        ElseIf Left$(varFormat, 5) = "round" Then
            On Error Resume Next
                decimalsCount = CLng(Mid$(varFormat, 6))
            On Error GoTo 0
            scope = CRFDecimal
        ElseIf Left$(varFormat, 10) = "percentage" Then
            On Error Resume Next
                decimalsCount = CLng(Mid$(varFormat, 11))
            On Error GoTo 0
            scope = CRFPercentage
        Else
            scope = CRFFormatCustom
        End If
    End Select

    If scope = 0 Then Exit Sub

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat valueCell, scope:=scope, nbDecimals:=decimalsCount
    On Error GoTo 0
End Sub

Private Sub WriteChoiceColumns(ByVal context As ILLVarContext, _
                               ByVal layout As ILLVarLayout)
    Dim controlType As String
    Dim categories As BetterArray
    Dim design As ILLFormat
    Dim choiceHeaderFlag As String
    Dim headerCell As Range
    Dim valueCell As Range
    Dim idx As Long

    controlType = LCase$(Trim$(context.ValueOf("control")))
    If controlType <> "choice_manual" Then Exit Sub

    Set categories = context.Categories(context.VariableName)
    If categories Is Nothing Then Exit Sub
    If categories.Length = 0 Then Exit Sub

    choiceHeaderFlag = LCase$(Trim$(context.ValueOf("crf choices")))
    Set design = context.Design
    Set valueCell = layout.ValueCell

    If valueCell Is Nothing Then Exit Sub

    If choiceHeaderFlag = "yes" Then
        Set headerCell = valueCell.Offset(-2, 1)
        For idx = categories.LowerBound To categories.UpperBound
            headerCell.Value = categories.Item(idx)
            If Not design Is Nothing Then
                On Error Resume Next
                    design.ApplyFormat headerCell, scope:=CRFChoiceHeader
                On Error GoTo 0
            End If
            Set headerCell = headerCell.Offset(0, 2)
        Next idx
    End If

    Set headerCell = valueCell.Offset(0, 1)
    For idx = categories.LowerBound To categories.UpperBound
        If Not design Is Nothing Then
            On Error Resume Next
                design.ApplyFormat headerCell, scope:=CRFChoiceValues
            On Error GoTo 0
        End If
        Set headerCell = headerCell.Offset(0, 2)
    Next idx
End Sub

'@section ILLVarWriterHooks implementation
'===============================================================================
Private Sub ILLVarWriterHooks_ConfigureLayout(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout)
    Dim crfSheet As Worksheet
    Dim anchorRow As Long

    If layout Is Nothing Then Exit Sub

    Set crfSheet = EnsureCRFWorksheet(context)
    anchorRow = ResolveCRFRow(context)
    layout.Configure crfSheet, anchorRow
End Sub

Private Sub ILLVarWriterHooks_AfterLabel(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout, _
                                         ByVal labelCell As Range, _
                                         ByVal mainLabelLength As Long)

    Dim hiddenRange As Range
    Dim tableName As String
    Dim anchorRow As Long

    ApplyLabelFormatting context, labelCell, mainLabelLength

    Set hiddenRange = layout.HiddenRange
    If Not hiddenRange Is Nothing Then
        hiddenRange.Hidden = ShouldHideRow(context.ValueOf("status"), context.ValueOf("control"))
    End If

    If Not labelCell Is Nothing Then
        On Error Resume Next
            tableName = context.ValueOf("table name")
            anchorRow = ResolveCRFRow(context)
            If anchorRow = 6 Then
                labelCell.Name = tableName & "_CRFSTART"
            End If
        On Error GoTo 0
    End If
End Sub

Private Sub ILLVarWriterHooks_AfterType(ByVal context As ILLVarContext, _
                                        ByVal layout As ILLVarLayout, _
                                        ByVal valueCell As Range)
    ApplyCRFTypeFormat context, valueCell
End Sub

Private Sub ILLVarWriterHooks_WriteFormulas(ByVal context As ILLVarContext, _
                                            ByVal layout As ILLVarLayout)
    'CRF variables do not inject specific formulas beyond choice metadata.
End Sub

Private Sub ILLVarWriterHooks_WriteChoices(ByVal context As ILLVarContext, _
                                           ByVal layout As ILLVarLayout)
    WriteChoiceColumns context, layout
End Sub

Private Sub ILLVarWriterHooks_AfterValidation(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout, _
                                              ByVal valueCell As Range)
    'CRF variables do not require additional validation adjustments.
End Sub

Private Sub ILLVarWriterHooks_AfterWrite(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout)
    'No finalisation required for CRF variables.
End Sub
