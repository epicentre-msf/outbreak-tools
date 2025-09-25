VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Reusable writer orchestrating shared linelist variable steps")
'@IgnoreModule UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriter

Private Type TLLVarWriter
    context As ILLVarContext
    layout As ILLVarLayout
    hooks As ILLVarWriterHooks
    isReady As Boolean
    layoutConfigured As Boolean
End Type

Private this As TLLVarWriter
Private Const CLASSNAME As String = "LLVarWriterBase"

'@section Lifecycle Helpers
'===============================================================================
Private Sub ResetInternal()
    Set this.context = Nothing
    Set this.layout = Nothing
    Set this.hooks = Nothing
    this.isReady = False
    this.layoutConfigured = False
End Sub

Private Sub EnsureReady()
    If Not this.isReady Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Writer must be initialised before invoking WriteVariable"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Initialisation
'===============================================================================
Private Sub ConfigureLayout()
    Dim labelCell As Range

    If this.layoutConfigured Then Exit Sub

    this.hooks.ConfigureLayout this.context, this.layout

    'Touch the value cell to confirm the layout is ready; this surfaces
    'misconfigured anchor indexes early in the workflow.
    On Error GoTo LayoutError
        Set labelCell = this.layout.ValueCell
    On Error GoTo 0

    this.layoutConfigured = True
    Exit Sub

LayoutError:
    ThrowError ProjectError.ObjectNotInitialized, _
               "Layout strategy failed to resolve the variable value cell"
End Sub

'@section Shared Operations
'===============================================================================
Private Sub WriteLabels()
    Dim mainLabel As String
    Dim subLabel As String
    Dim fullLabel As String
    Dim varName As String
    Dim varStatus As String
    Dim varNote As String
    Dim calculatedTag As String
    Dim mandatoryTag As String
    Dim controlType As String
    Dim translation As ITranslationObject
    Dim labelCell As Range
    Dim nameCell As Range
    Dim noteCell As Range
    Dim mainLabelLength As Long

    mainLabel = Trim$(this.context.ValueOf("main label"))
    subLabel = Trim$(this.context.ValueOf("sub label"))
    varName = this.context.ValueOf("variable name")
    varStatus = LCase$(Trim$(this.context.ValueOf("status")))
    varNote = this.context.ValueOf("note")
    controlType = LCase$(Trim$(this.context.ValueOf("control")))

    Set translation = this.context.Translation

    If Not translation Is Nothing Then
        calculatedTag = translation.TranslatedValue("MSG_Calculated")
        mandatoryTag = translation.TranslatedValue("MSG_Mandatory")
    End If

    If (controlType = "formula") Or (controlType = "case_when") Or _
       (controlType = "choice_formula") Then
        If LenB(calculatedTag) > 0 Then
            If LenB(subLabel) > 0 Then
                subLabel = subLabel & vbLf & calculatedTag
            Else
                subLabel = calculatedTag
            End If
        End If
    End If

    If LenB(subLabel) > 0 Then
        fullLabel = mainLabel & vbLf & subLabel
    Else
        fullLabel = mainLabel
    End If

    mainLabelLength = Len(mainLabel)

    If varStatus = "mandatory" Then
        If LenB(mandatoryTag) > 0 Then
            If LenB(varNote) > 0 Then
                varNote = mandatoryTag & vbLf & varNote
            Else
                varNote = mandatoryTag
            End If
        End If
    End If

    Set labelCell = this.layout.LabelCell
    Set nameCell = this.layout.NameCell
    Set noteCell = this.layout.NoteTarget

    If Not labelCell Is Nothing Then
        labelCell.Value = fullLabel
    End If

    If Not nameCell Is Nothing Then
        nameCell.Value = varName
    End If

    If Not noteCell Is Nothing Then
        On Error Resume Next
            If Not noteCell.Comment Is Nothing Then
                noteCell.Comment.Delete
            End If
        On Error GoTo 0

        If LenB(varNote) > 0 Then
            On Error Resume Next
                noteCell.AddComment Text:=varNote
                noteCell.Comment.Visible = False
            On Error GoTo 0
        End If
    End If

    this.hooks.AfterLabel this.context, this.layout, labelCell, mainLabelLength
End Sub

Private Sub ApplyValueType()
    Dim varType As String
    Dim varFormat As String
    Dim targetCell As Range
    Dim printedCell As Range
    Dim design As ILLFormat
    Dim numberFormatText As String
    Dim nbDecimals As Long

    varType = LCase$(Trim$(this.context.ValueOf("variable type")))
    varFormat = LCase$(Trim$(this.context.ValueOf("variable format")))

    Set targetCell = this.layout.ValueCell
    Set printedCell = this.layout.PrintedValueCell
    Set design = this.context.Design

    Select Case varType
    Case "text"
        If Not targetCell Is Nothing Then targetCell.NumberFormat = "@"
        If Not printedCell Is Nothing Then printedCell.NumberFormat = "@"

    Case "date"
        If LenB(varFormat) = 0 Then varFormat = "d-mmm-yyyy"
        If Not targetCell Is Nothing Then targetCell.NumberFormat = varFormat
        If Not printedCell Is Nothing Then printedCell.NumberFormat = varFormat

    Case "decimal", "integer"
        If LenB(varFormat) = 0 Then
            numberFormatText = IIf(varType = "decimal", _
                                   "#,##0.00;-#,##0.00;0.00;@", _
                                   "###0;-###0;0;@")
        ElseIf varFormat = "integer" Then
            numberFormatText = "###0;-###0;0;@"
        ElseIf varFormat = "round0" Then
            numberFormatText = "#,##0;-#,##0;0;@"
        ElseIf Left$(varFormat, 5) = "round" Then
            nbDecimals = 0
            On Error Resume Next
                nbDecimals = CLng(Mid$(varFormat, 6))
            On Error GoTo 0
            numberFormatText = "#,###0." & String(nbDecimals, "0") & _
                               ";-#,###0." & String(nbDecimals, "0") & _
                               ";0." & String(nbDecimals, "0") & ";@"
        ElseIf varFormat = "percentage0" Then
            numberFormatText = "0%"
        ElseIf Left$(varFormat, 10) = "percentage" Then
            nbDecimals = 0
            On Error Resume Next
                nbDecimals = CLng(Mid$(varFormat, 11))
            On Error GoTo 0
            numberFormatText = "0." & String(nbDecimals, "0") & "%"
        Else
            numberFormatText = varFormat
        End If

        If Not targetCell Is Nothing Then targetCell.NumberFormat = numberFormatText
        If Not printedCell Is Nothing Then printedCell.NumberFormat = numberFormatText

    Case Else
        'Leave the cell using workbook defaults
    End Select

    If Not design Is Nothing Then
        If Not targetCell Is Nothing Then design.ApplyFormat targetCell, FormatScopeLinelistHiddenCell
    End If

    this.hooks.AfterType this.context, this.layout, targetCell
End Sub

Private Sub ApplyValidation()
    Dim varType As String
    Dim varMin As String
    Dim varMax As String
    Dim varAlert As String
    Dim varMessage As String
    Dim validationType As Long
    Dim validationAlert As Long
    Dim targetCell As Range

    varType = LCase$(Trim$(this.context.ValueOf("variable type")))
    varMin = Trim$(this.context.ValueOf("min"))
    varMax = Trim$(this.context.ValueOf("max"))
    varAlert = LCase$(Trim$(this.context.ValueOf("alert")))
    varMessage = this.context.ValueOf("message")

    Set targetCell = this.layout.ValueCell

    If targetCell Is Nothing Then Exit Sub
    If (LenB(varMin) = 0 And LenB(varMax) = 0) Then GoTo NotifyHooks
    If varType = "text" Then GoTo NotifyHooks

    Select Case varType
    Case "integer"
        validationType = xlValidateWholeNumber
    Case "decimal"
        validationType = xlValidateDecimal
    Case "date"
        validationType = xlValidateDate
    Case Else
        GoTo NotifyHooks
    End Select

    Select Case varAlert
    Case "error"
        validationAlert = xlValidAlertStop
    Case "warning"
        validationAlert = xlValidAlertWarning
    Case Else
        validationAlert = xlValidAlertInformation
    End Select

    On Error Resume Next
        targetCell.Validation.Delete
    On Error GoTo 0

    On Error Resume Next
        Select Case True
        Case (LenB(varMin) > 0) And (LenB(varMax) = 0)
            targetCell.Validation.Add Type:=validationType, AlertStyle:=validationAlert, _
                Operator:=xlGreaterEqual, Formula1:=varMin

        Case (LenB(varMin) = 0) And (LenB(varMax) > 0)
            targetCell.Validation.Add Type:=validationType, AlertStyle:=validationAlert, _
                Operator:=xlLessEqual, Formula1:=varMax

        Case (LenB(varMin) > 0) And (LenB(varMax) > 0)
            targetCell.Validation.Add Type:=validationType, AlertStyle:=validationAlert, _
                Operator:=xlBetween, Formula1:=varMin, Formula2:=varMax
        End Select

        With targetCell.Validation
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = vbNullString
            .InputMessage = vbNullString
            .ErrorTitle = vbNullString
            .ErrorMessage = varMessage
            .ShowInput = True
            .ShowError = True
        End With
    On Error GoTo 0

NotifyHooks:
    this.hooks.AfterValidation this.context, this.layout, targetCell
End Sub

'@section Template Sequence
'===============================================================================
Private Sub RunWriteSequence()
    ConfigureLayout
    WriteLabels
    ApplyValueType
    this.hooks.WriteFormulas this.context, this.layout
    this.hooks.WriteChoices this.context, this.layout
    ApplyValidation
    this.hooks.AfterWrite this.context, this.layout
End Sub

'@section Interface Implementation
'===============================================================================
Private Sub ILLVarWriter_Initialise(ByVal context As ILLVarContext, _
                                    ByVal layout As ILLVarLayout, _
                                    ByVal hooks As ILLVarWriterHooks)

    If context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Context instance is required"
    End If

    If layout Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Layout strategy is required"
    End If

    If hooks Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Writer hooks instance is required"
    End If

    ResetInternal

    Set this.context = context
    Set this.layout = layout
    Set this.hooks = hooks
    this.isReady = True
End Sub

Private Sub ILLVarWriter_Reset()
    ResetInternal
End Sub

Private Sub ILLVarWriter_WriteVariable()
    EnsureReady
    RunWriteSequence
End Sub

