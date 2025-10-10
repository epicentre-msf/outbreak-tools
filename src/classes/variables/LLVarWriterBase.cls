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
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriter

Private Type TLLVarWriter
    context As ILLVarContext
    layout As ILLVarLayout
    hooks As ILLVarWriterHooks
    isReady As Boolean
    layoutConfigured As Boolean
End Type

Private Type TValidationOperand
    hasValue As Boolean
    useFormula As Boolean
    typedValue As Variant
    formulaText As String
    osFormulaText As String
    rawText As String
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

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
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
    Dim valueCell As Range
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
    Set valueCell = this.layout.ValueCell
    Set nameCell = this.layout.NameCell
    Set noteCell = this.layout.NoteTarget

    If Not labelCell Is Nothing Then
        labelCell.Value = fullLabel
    End If

    If Not valueCell Is Nothing Then
        ApplyRangeName valueCell, varName
    End If

    If Not nameCell Is Nothing Then
        'Vertical layouts reuse the value cell as the name cell; keep the cell empty
        'while still allowing horizontal layouts to display the variable identifier.
        If Not valueCell Is Nothing Then
            If IsSameCell(nameCell, valueCell) Then
                nameCell.Value = vbNullString
            Else
                nameCell.Value = varName
            End If
        Else
            nameCell.Value = varName
        End If
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

Private Sub ApplyRangeName(ByVal target As Range, ByVal rangeName As String)
    '@description Assign a workbook-level name to the supplied cell, mirroring the legacy writers.
    If target Is Nothing Then Exit Sub
    If LenB(rangeName) = 0 Then Exit Sub

    On Error GoTo ApplyNameFailed
        target.Name = rangeName
    On Error GoTo 0
    Exit Sub

ApplyNameFailed:
    'Invalid names should not block the writer: clear the error and proceed without naming.
    Err.Clear
    On Error GoTo 0
End Sub

Private Function IsSameCell(ByVal first As Range, ByVal second As Range) As Boolean
    '@description Helper to determine if two range handles refer to the same cell.
    Dim firstSheet As Worksheet
    Dim secondSheet As Worksheet
    Dim sameSheet As Boolean

    If first Is Nothing Then Exit Function
    If second Is Nothing Then Exit Function

    Set firstSheet = first.Worksheet
    Set secondSheet = second.Worksheet

    If firstSheet Is Nothing Then Exit Function
    If secondSheet Is Nothing Then Exit Function

    sameSheet = (firstSheet Is secondSheet)

    If sameSheet Then
        IsSameCell = (StrComp(first.Address(False, False), second.Address(False, False), vbTextCompare) = 0)
    ElseIf firstSheet.Parent Is secondSheet.Parent Then
        IsSameCell = (StrComp(first.Address(False, False, xlA1, True), _
                              second.Address(False, False, xlA1, True), _
                              vbTextCompare) = 0)
    End If
End Function

Private Sub ApplyNumberFormatTo(ByVal target As Range, _
                                ByVal formatText As String, _
                                Optional ByVal allowFailure As Boolean = False)
    '@description Wrapper that assigns NumberFormat while tolerating custom formats when requested.
    If target Is Nothing Then Exit Sub
    If LenB(formatText) = 0 Then Exit Sub

    If allowFailure Then
        On Error Resume Next
            target.NumberFormat = formatText
        On Error GoTo 0
    Else
        target.NumberFormat = formatText
    End If
End Sub

Private Function ChooseDateFormat(ByVal rawFormat As String, ByVal defaultFormat As String) As String
    '@description Preserve author supplied formats for dates when available, fall back to defaults otherwise.
    Dim cleanedFormat As String

    cleanedFormat = Trim$(rawFormat)

    If LenB(cleanedFormat) > 0 Then
        ChooseDateFormat = cleanedFormat
    Else
        ChooseDateFormat = defaultFormat
    End If
End Function

Private Function ResolveValidationType(ByVal varType As String) As Long
    Select Case varType
    Case "integer"
        ResolveValidationType = xlValidateWholeNumber
    Case "decimal"
        ResolveValidationType = xlValidateDecimal
    Case "date"
        ResolveValidationType = xlValidateDate
    Case Else
        ResolveValidationType = 0
    End Select
End Function

Private Function ResolveValidationAlert(ByVal varAlert As String) As Long
    Select Case varAlert
    Case "error"
        ResolveValidationAlert = xlValidAlertStop
    Case "warning"
        ResolveValidationAlert = xlValidAlertWarning
    Case Else
        ResolveValidationAlert = xlValidAlertInformation
    End Select
End Function

Private Sub ApplyValidationSingle(ByVal validationObj As Validation, _
                                  ByVal validationType As Long, _
                                  ByVal validationAlert As Long, _
                                  ByVal operatorType As Long, _
                                  ByRef operand As TValidationOperand)
    If validationObj Is Nothing Then Exit Sub
    If Not operand.hasValue Then Exit Sub

    If operand.useFormula Then
        AddValidationInternal validationObj, validationType, validationAlert, operatorType, _
                               DefaultFormulaValue(operand.osFormulaText, operand.formulaText)
        AddValidationInternal validationObj, validationType, validationAlert, operatorType, _
                               DefaultFormulaValue(operand.formulaText, operand.osFormulaText)
    Else
        AddValidationInternal validationObj, validationType, validationAlert, operatorType, _
                               operand.typedValue
    End If
End Sub

Private Sub ApplyValidationBetween(ByVal validationObj As Validation, _
                                   ByVal validationType As Long, _
                                   ByVal validationAlert As Long, _
                                   ByRef minOperand As TValidationOperand, _
                                   ByRef maxOperand As TValidationOperand)
    Dim useFormula As Boolean
    Dim firstFormula1 As Variant
    Dim firstFormula2 As Variant
    Dim secondFormula1 As Variant
    Dim secondFormula2 As Variant

    If validationObj Is Nothing Then Exit Sub
    If Not minOperand.hasValue Or Not maxOperand.hasValue Then Exit Sub

    useFormula = minOperand.useFormula Or maxOperand.useFormula

    If useFormula Then
        firstFormula1 = DefaultFormulaValue(minOperand.osFormulaText, minOperand.formulaText)
        firstFormula2 = DefaultFormulaValue(maxOperand.osFormulaText, maxOperand.formulaText)
        secondFormula1 = DefaultFormulaValue(minOperand.formulaText, minOperand.osFormulaText)
        secondFormula2 = DefaultFormulaValue(maxOperand.formulaText, maxOperand.osFormulaText)

        AddValidationInternal validationObj, validationType, validationAlert, xlBetween, _
                               firstFormula1, firstFormula2
        AddValidationInternal validationObj, validationType, validationAlert, xlBetween, _
                               secondFormula1, secondFormula2
    Else
        AddValidationInternal validationObj, validationType, validationAlert, xlBetween, _
                               minOperand.typedValue, maxOperand.typedValue
    End If
End Sub

Private Function DefaultFormulaValue(ByVal primaryValue As String, _
                                     ByVal fallbackValue As String) As String
    If LenB(primaryValue) > 0 Then
        DefaultFormulaValue = primaryValue
    Else
        DefaultFormulaValue = fallbackValue
    End If
End Function

Private Sub AddValidationInternal(ByVal validationObj As Validation, _
                                  ByVal validationType As Long, _
                                  ByVal validationAlert As Long, _
                                  ByVal operatorType As Long, _
                                  ByVal formula1 As Variant, _
                                  Optional ByVal formula2 As Variant)
    If validationObj Is Nothing Then Exit Sub
    If IsMissing(formula2) Then
        On Error Resume Next
            validationObj.Add Type:=validationType, AlertStyle:=validationAlert, _
                              Operator:=operatorType, Formula1:=formula1
        On Error GoTo 0
    Else
        On Error Resume Next
            validationObj.Add Type:=validationType, AlertStyle:=validationAlert, _
                              Operator:=operatorType, Formula1:=formula1, Formula2:=formula2
        On Error GoTo 0
    End If
End Sub

Private Function BuildValidationOperand(ByVal rawValue As String, _
                                        ByVal validationType As Long, _
                                        ByVal hostCell As Range) As TValidationOperand
    Dim operand As TValidationOperand
    Dim trimmedValue As String
    Dim specs As ILinelistSpecs
    Dim dict As ILLdictionary
    Dim formulaData As IFormulaData
    Dim formulaEngine As IFormulas
    Dim parsedFormula As String
    Dim evaluationResult As Variant
    Dim isGroupedFormula As Boolean
    Dim success As Boolean

    trimmedValue = Trim$(rawValue)
    If LenB(trimmedValue) = 0 Then
        BuildValidationOperand = operand
        Exit Function
    End If

    operand.hasValue = True
    operand.rawText = trimmedValue

    Set specs = SafeSpecs()
    If Not specs Is Nothing Then
        On Error Resume Next
            Set dict = specs.Dictionary
            Set formulaData = specs.FormulaDataObject
        On Error GoTo 0

        If (Not dict Is Nothing) And (Not formulaData Is Nothing) Then
            On Error Resume Next
                Set formulaEngine = Formulas.Create(dict, formulaData, trimmedValue)
            On Error GoTo 0

            If Not formulaEngine Is Nothing Then
                If formulaEngine.Valid("linelist") Then
                    isGroupedFormula = (StrComp(formulaEngine.IsGrouped, "Yes", vbTextCompare) = 0)
                    parsedFormula = formulaEngine.ParsedLinelistFormula()

                    If LenB(parsedFormula) > 0 Then
                        If HasVariableReferences(formulaEngine) Then
                            operand.useFormula = True
                            operand.formulaText = "= " & parsedFormula
                            operand.osFormulaText = ConvertToOSFormula(hostCell, parsedFormula, isGroupedFormula)
                            If LenB(operand.osFormulaText) = 0 Then
                                operand.osFormulaText = operand.formulaText
                            End If
                            success = True
                        Else
                            evaluationResult = EvaluateExpression(parsedFormula)
                            If ShouldUseTypedValue(evaluationResult) Then
                                operand.useFormula = False
                                operand.typedValue = CoerceToValidationType(evaluationResult, validationType, trimmedValue)
                                success = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If Not success Then
        operand.useFormula = False
        operand.typedValue = ConvertLiteralValue(trimmedValue, validationType, trimmedValue)
    End If

    BuildValidationOperand = operand
End Function

Private Function SafeSpecs() As ILinelistSpecs
    On Error Resume Next
        Set SafeSpecs = this.context.Specs
    On Error GoTo 0
End Function

Private Function ConvertLiteralValue(ByVal expression As String, _
                                     ByVal validationType As Long, _
                                     ByVal fallbackText As String) As Variant
    Dim evaluationSource As String
    Dim evaluationResult As Variant

    evaluationSource = RemoveLeadingEquals(expression)

    If LenB(evaluationSource) = 0 Then
        ConvertLiteralValue = fallbackText
        Exit Function
    End If

    evaluationResult = EvaluateExpression(evaluationSource)

    If ShouldUseTypedValue(evaluationResult) Then
        ConvertLiteralValue = CoerceToValidationType(evaluationResult, validationType, fallbackText)
    Else
        ConvertLiteralValue = CoerceToValidationType(evaluationSource, validationType, fallbackText)
    End If
End Function

Private Function CoerceToValidationType(ByVal value As Variant, _
                                        ByVal validationType As Long, _
                                        ByVal fallbackText As String) As Variant
    On Error GoTo ConversionFailed

    Select Case validationType
    Case xlValidateWholeNumber
        CoerceToValidationType = CLng(value)
    Case xlValidateDecimal
        CoerceToValidationType = CDbl(value)
    Case xlValidateDate
        CoerceToValidationType = CDate(value)
    Case Else
        CoerceToValidationType = value
    End Select
    Exit Function

ConversionFailed:
    Err.Clear
    CoerceToValidationType = fallbackText
End Function

Private Function RemoveLeadingEquals(ByVal expression As String) As String
    Dim working As String

    working = Trim$(expression)

    Do While LenB(working) > 0 And Left$(working, 1) = "="
        working = Trim$(Mid$(working, 2))
    Loop

    RemoveLeadingEquals = working
End Function

Private Function EvaluateExpression(ByVal expression As String) As Variant
    Dim result As Variant

    If LenB(expression) = 0 Then Exit Function

    On Error Resume Next
        result = Application.Evaluate("=" & expression)
    On Error GoTo 0

    EvaluateExpression = result
End Function

Private Function ShouldUseTypedValue(ByVal evaluationResult As Variant) As Boolean
    If IsError(evaluationResult) Then Exit Function

    Select Case VarType(evaluationResult)
    Case vbEmpty, vbNull, vbError, vbString, vbBoolean
        ShouldUseTypedValue = False
    Case Else
        ShouldUseTypedValue = True
    End Select
End Function

Private Function ConvertToOSFormula(ByVal hostCell As Range, _
                                    ByVal parsedFormula As String, _
                                    Optional ByVal useArrayFormula As Boolean = False) As String
    Dim testCell As Range
    Dim previousFormula As String
    Dim previousValue As Variant
    Dim previousHasFormula As Boolean
    Dim previousArrayState As Boolean

    If LenB(parsedFormula) = 0 Then Exit Function

    If hostCell Is Nothing Then
        ConvertToOSFormula = "= " & parsedFormula
        Exit Function
    End If

    Set testCell = hostCell.Worksheet.Cells(1, 1)
    previousHasFormula = testCell.HasFormula
    previousArrayState = testCell.HasArray
    previousFormula = testCell.Formula
    previousValue = testCell.Value

    On Error GoTo RestoreFormula
        If useArrayFormula Then
            testCell.FormulaArray = "=" & parsedFormula
        Else
            testCell.Formula = "=" & parsedFormula
        End If
        ConvertToOSFormula = testCell.FormulaLocal
RestoreFormula:
    On Error Resume Next
        If previousHasFormula Then
            If previousArrayState Then
                If LenB(previousFormula) > 0 Then
                    testCell.FormulaArray = previousFormula
                Else
                    testCell.ClearContents
                End If
            Else
                If LenB(previousFormula) > 0 Then
                    testCell.Formula = previousFormula
                Else
                    testCell.ClearContents
                End If
            End If
        Else
            testCell.ClearContents
            On Error Resume Next
                testCell.Value = previousValue
            On Error GoTo 0
        End If
    On Error GoTo 0

    If LenB(ConvertToOSFormula) = 0 Then
        ConvertToOSFormula = "= " & parsedFormula
    End If
End Function

Private Function HasVariableReferences(ByVal formulaEngine As IFormulas) As Boolean
    Dim referenced As BetterArray

    If formulaEngine Is Nothing Then Exit Function

    On Error Resume Next
        Set referenced = formulaEngine.Varlists("linelist")
    On Error GoTo 0

    If referenced Is Nothing Then Exit Function

    HasVariableReferences = (referenced.Length > 0)
End Function

Private Function CombineRange(ByVal first As Range, ByVal second As Range) As Range
    If first Is Nothing Then
        If second Is Nothing Then Exit Function
        Set CombineRange = second
    ElseIf second Is Nothing Then
        Set CombineRange = first
    Else
        On Error Resume Next
            Set CombineRange = Union(first, second)
        On Error GoTo 0

        If CombineRange Is Nothing Then
            Set CombineRange = first
        End If
    End If
End Function

Private Sub UnlockCell(ByVal target As Range)
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        target.Locked = False
    On Error GoTo 0
End Sub

Private Sub ApplyHiddenFormat(ByVal design As ILLFormat, ByVal target As Range)
    If design Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat target, LinelistHiddenCell
    On Error GoTo 0
End Sub

Private Sub ApplyValueType()
    Dim varType As String
    Dim varFormatSource As String
    Dim varFormat As String
    Dim targetCell As Range
    Dim printedCell As Range
    Dim design As ILLFormat
    Dim numberFormatText As String
    Dim nbDecimals As Long
    Dim dateFormatText As String
    Dim useCustomFormat As Boolean

    varType = LCase$(Trim$(this.context.ValueOf("variable type")))
    varFormatSource = Trim$(this.context.ValueOf("variable format"))
    varFormat = LCase$(varFormatSource)

    Set targetCell = this.layout.ValueCell
    Set printedCell = this.layout.PrintedValueCell
    Set design = this.context.Design

    Select Case varType
    Case "text"
        ApplyNumberFormatTo targetCell, "@"
        ApplyNumberFormatTo printedCell, "@"

    Case "date"
        If LenB(varFormat) = 0 Then varFormat = "d-mmm-yyyy"
        dateFormatText = ChooseDateFormat(varFormatSource, varFormat)
        ApplyNumberFormatTo targetCell, dateFormatText, True
        ApplyNumberFormatTo printedCell, dateFormatText, True

    Case "decimal", "integer"
        useCustomFormat = False

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
        ElseIf varFormat = "euros" Then
            numberFormatText = "#,##0.00 [$€-x-euro1];-#,##0.00 [$€-x-euro1];0.00 [$€-x-euro1];@"
        ElseIf varFormat = "dollars" Then
            numberFormatText = "[$$-en-US] #,##0.00;[$$-en-US] -#,##0.00;[$$-en-US] 0.00;@"
        Else
            numberFormatText = IIf(LenB(varFormatSource) > 0, varFormatSource, varFormat)
            useCustomFormat = True
        End If

        ApplyNumberFormatTo targetCell, numberFormatText, useCustomFormat
        ApplyNumberFormatTo printedCell, numberFormatText, useCustomFormat

    Case Else
        'Leave the cell using workbook defaults
    End Select

    If Not design Is Nothing Then
        If Not targetCell Is Nothing Then design.ApplyFormat targetCell, LinelistHiddenCell
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
    Dim validationObj As Validation
    Dim minOperand As TValidationOperand
    Dim maxOperand As TValidationOperand
    Dim hasConstraints As Boolean

    varType = LCase$(Trim$(this.context.ValueOf("variable type")))
    varMin = this.context.ValueOf("min")
    varMax = this.context.ValueOf("max")
    varAlert = LCase$(Trim$(this.context.ValueOf("alert")))
    varMessage = this.context.ValueOf("message")

    Set targetCell = this.layout.ValueCell

    If targetCell Is Nothing Then Exit Sub

    hasConstraints = (LenB(Trim$(varMin)) > 0) Or (LenB(Trim$(varMax)) > 0)
    If Not hasConstraints Then GoTo NotifyHooks
    If varType = "text" Then GoTo NotifyHooks

    validationType = ResolveValidationType(varType)
    If validationType = 0 Then GoTo NotifyHooks

    validationAlert = ResolveValidationAlert(varAlert)

    On Error Resume Next
        targetCell.Validation.Delete
    On Error GoTo 0

    minOperand = BuildValidationOperand(varMin, validationType, targetCell)
    maxOperand = BuildValidationOperand(varMax, validationType, targetCell)

    Set validationObj = targetCell.Validation

    Select Case True
    Case minOperand.hasValue And Not maxOperand.hasValue
        ApplyValidationSingle validationObj, validationType, validationAlert, _
                              xlGreaterEqual, minOperand

    Case (Not minOperand.hasValue) And maxOperand.hasValue
        ApplyValidationSingle validationObj, validationType, validationAlert, _
                              xlLessEqual, maxOperand

    Case minOperand.hasValue And maxOperand.hasValue
        ApplyValidationBetween validationObj, validationType, validationAlert, _
                               minOperand, maxOperand

    Case Else
        GoTo NotifyHooks
    End Select

    With validationObj
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorTitle = vbNullString
        .ErrorMessage = varMessage
        .ShowInput = True
        .ShowError = True
    End With

NotifyHooks:
    this.hooks.AfterValidation this.context, this.layout, targetCell
End Sub

Private Sub ApplyProtection()
    Dim valueCell As Range
    Dim controlCell As Range
    Dim autoOriginCell As Range
    Dim metadataRange As Range
    Dim design As ILLFormat

    Set valueCell = this.layout.ValueCell
    Set controlCell = this.layout.ControlCell
    Set autoOriginCell = this.layout.AutoOriginCell
    Set metadataRange = CombineRange(controlCell, autoOriginCell)
    Set design = this.context.Design

    UnlockCell valueCell
    ApplyHiddenFormat design, controlCell
    ApplyHiddenFormat design, autoOriginCell

    this.hooks.ConfigureProtection this.context, this.layout, valueCell, metadataRange
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
    ApplyProtection
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

