VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterHooksHorizontal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Default hooks for horizontal variable writer")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriterHooks

Private Const START_COLUMN_INDEX As Long = 1

Private Sub ILLVarWriterHooks_AfterLabel(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout, _
                                         ByVal labelCell As Range, _
                                         ByVal mainLabelLength As Long)
    ApplyMainLabelFormatting context, layout, labelCell, mainLabelLength
End Sub

Private Sub ILLVarWriterHooks_AfterType(ByVal context As ILLVarContext, _
                                        ByVal layout As ILLVarLayout, _
                                        ByVal valueCell As Range)
    'No additional behaviour required
End Sub

Private Sub ILLVarWriterHooks_AfterValidation(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout, _
                                              ByVal valueCell As Range)
    ApplyFormattingCondition context, valueCell
    ApplyUniqueCondition context, valueCell
    ApplyGeoWarning context, valueCell
End Sub

Private Sub ILLVarWriterHooks_AfterWrite(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout)
    ApplyAnchors context, layout
    SynchronisePrintedHeader context, layout, Len(context.ValueOf("main label"))
    PopulateHiddenMetadata context, layout
    ApplyGeoStyling context, layout
    ApplyEditableLabel context, layout
    ApplyRegisterBookRules context, layout
End Sub

Private Sub ILLVarWriterHooks_ConfigureLayout(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout)
    Dim anchorColumn As Long
    Dim sheet As Worksheet
    Dim printSheet As Worksheet

    anchorColumn = CLng(context.ValueOf("column index"))
    Set sheet = context.Worksheet
    Set printSheet = Nothing

    On Error Resume Next
        Set printSheet = context.PrintWorksheet
    On Error GoTo 0

    layout.Configure sheet, anchorColumn, printSheet
End Sub

Private Sub ILLVarWriterHooks_WriteChoices(ByVal context As ILLVarContext, _
                                           ByVal layout As ILLVarLayout)
    Dim controlType As String

    controlType = NormaliseControlType(context.ValueOf("control"))

    Select Case controlType
    Case "choice_manual", "choice_multiple"
        HandleStandardChoices context, layout, controlType

    Case "choice_custom"
        HandleCustomChoices context, layout

    Case "list_auto"
        HandleListAuto context, layout

    Case "geo1"
        HandleGeoChoices context, layout
    End Select
End Sub

Private Sub ILLVarWriterHooks_WriteFormulas(ByVal context As ILLVarContext, _
                                            ByVal layout As ILLVarLayout)
    Dim controlType As String
    Dim controlDetails As String
    Dim valueCell As Range
    Dim labelCell As Range
    Dim design As ILLFormat
    Dim dict As ILLdictionary
    Dim specs As ILinelistSpecs
    Dim formulaData As IFormulaData
    Dim formulaEngine As IFormulas
    Dim parsedFormula As String
    Dim isGroupedFormula As Boolean

    controlType = NormaliseControlType(context.ValueOf("control"))
    If Not IsCalculatedControl(controlType) Then Exit Sub

    controlDetails = Trim$(context.ValueOf("control details"))
    If LenB(controlDetails) = 0 Then Exit Sub

    Set valueCell = layout.ValueCell
    Set labelCell = layout.LabelCell
    If valueCell Is Nothing Then Exit Sub

    Set dict = SafeDictionary(context)
    Set specs = SafeSpecs(context)
    If dict Is Nothing Then Exit Sub
    If specs Is Nothing Then Exit Sub

    On Error Resume Next
        Set formulaData = specs.FormulaDataObject
    On Error GoTo 0
    If formulaData Is Nothing Then Exit Sub

    On Error Resume Next
        Set formulaEngine = Formulas.Create(dict, formulaData, controlDetails)
    On Error GoTo 0
    If formulaEngine Is Nothing Then Exit Sub
    If Not formulaEngine.Valid("linelist") Then Exit Sub

    parsedFormula = formulaEngine.ParsedLinelistFormula()
    isGroupedFormula = (StrComp(formulaEngine.IsGrouped, "Yes", vbTextCompare) = 0)
    If LenB(parsedFormula) = 0 Then Exit Sub

    If Not ApplyParsedFormula(valueCell, parsedFormula, isGroupedFormula) Then
        ApplyFormulaFailure valueCell
        Exit Sub
    End If

    valueCell.Locked = True
    valueCell.FormulaHidden = True

    Set design = context.Design
    ApplyCalculatedCellFormat design, valueCell
    ApplyCalculatedHeaderFormat design, labelCell, layout.PrintedValueCell
End Sub

Private Sub ILLVarWriterHooks_ConfigureProtection(ByVal context As ILLVarContext, _
                                                  ByVal layout As ILLVarLayout, _
                                                  ByVal valueCell As Range, _
                                                  ByVal metadataRange As Range)
    'Default implementation does not adjust protection
End Sub

Private Sub ApplyMainLabelFormatting(ByVal context As ILLVarContext, _
                                     ByVal layout As ILLVarLayout, _
                                     ByVal labelCell As Range, _
                                     ByVal mainLabelLength As Long)
    Dim design As ILLFormat
    Dim labelText As String

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    If Not labelCell Is Nothing Then
        labelText = CStr(labelCell.Value)
        If LenB(labelText) > 0 Then
            ApplyLabelScopes design, labelCell, mainLabelLength
        End If
    End If

    SynchronisePrintedHeader context, layout, mainLabelLength
End Sub

Private Sub SynchronisePrintedHeader(ByVal context As ILLVarContext, _
                                     ByVal layout As ILLVarLayout, _
                                     Optional ByVal mainLabelLength As Long = 0)
    Dim printedValue As Range
    Dim printedLabel As Range
    Dim printedName As Range
    Dim design As ILLFormat
    Dim sourceLabel As Range
    Dim labelText As String

    Set printedValue = layout.PrintedValueCell
    If printedValue Is Nothing Then Exit Sub

    Set printedLabel = printedValue.Offset(-2, 0)
    Set printedName = printedValue.Offset(-1, 0)
    Set sourceLabel = layout.LabelCell

    If Not sourceLabel Is Nothing Then
        labelText = CStr(sourceLabel.Value)
    Else
        labelText = vbNullString
    End If

    If LenB(labelText) > 0 Then
        printedLabel.Value = labelText
    End If

    printedName.Value = context.ValueOf("variable name")

    Set design = context.Design
    If Not design Is Nothing Then
        If LenB(labelText) > 0 Then
            ApplyLabelScopes design, printedLabel, mainLabelLength
        End If
    End If
End Sub

Private Sub ApplyAnchors(ByVal context As ILLVarContext, ByVal layout As ILLVarLayout)
    Dim columnIndex As Long
    Dim tableName As String
    Dim mainNameCell As Range
    Dim printedNameCell As Range

    columnIndex = CLng(Val(context.ValueOf("column index")))
    If columnIndex <> START_COLUMN_INDEX Then Exit Sub

    tableName = Trim$(context.ValueOf("table name"))
    If LenB(tableName) = 0 Then Exit Sub

    Set mainNameCell = layout.NameCell
    AssignWorkbookName mainNameCell, tableName & "_START"

    Set printedNameCell = PrintedNameCell(layout)
    If Not printedNameCell Is Nothing Then
        AssignWorkbookName printedNameCell, tableName & "_PRINTSTART"
    End If
End Sub

Private Sub ApplyLabelScopes(ByVal design As ILLFormat, _
                             ByVal labelCell As Range, _
                             ByVal mainLabelLength As Long)
    Dim totalLength As Long

    If design Is Nothing Then Exit Sub
    If labelCell Is Nothing Then Exit Sub
    If mainLabelLength <= 0 Then Exit Sub

    totalLength = Len(CStr(labelCell.Value))

    On Error Resume Next
        design.ApplyFormat labelCell.Characters(Start:=1, Length:=mainLabelLength), _
                           scope:=HListMainLab
        If totalLength > mainLabelLength Then
            design.ApplyFormat labelCell.Characters(Start:=mainLabelLength + 1), _
                               scope:=HListSublab
        End If
    On Error GoTo 0
End Sub

Private Function PrintedNameCell(ByVal layout As ILLVarLayout) As Range
    Dim printedValue As Range
    Set printedValue = layout.PrintedValueCell
    If printedValue Is Nothing Then Exit Function
    Set PrintedNameCell = printedValue.Offset(-1, 0)
End Function

Private Sub AssignWorkbookName(ByVal target As Range, ByVal rangeName As String)
    Dim hostWorkbook As Workbook

    If target Is Nothing Then Exit Sub
    If LenB(rangeName) = 0 Then Exit Sub

    Set hostWorkbook = target.Worksheet.Parent

    On Error Resume Next
        hostWorkbook.Names(rangeName).Delete
    On Error GoTo 0

    On Error Resume Next
        hostWorkbook.Names.Add Name:=rangeName, RefersTo:=target
    On Error GoTo 0
End Sub

Private Sub PopulateHiddenMetadata(ByVal context As ILLVarContext, _
                                   ByVal layout As ILLVarLayout)
    Dim controlValue As String
    Dim listOrigin As String
    Dim controlCell As Range
    Dim autoOriginCell As Range
    Dim design As ILLFormat

    controlValue = context.ValueOf("control")
    listOrigin = context.ValueOf("list auto")

    Set controlCell = layout.ControlCell
    Set autoOriginCell = layout.AutoOriginCell
    Set design = context.Design

    If Not controlCell Is Nothing Then
        controlCell.Value = controlValue
        ApplyHiddenFormat design, controlCell
    End If

    If Not autoOriginCell Is Nothing Then
        If LenB(listOrigin) > 0 Then
            autoOriginCell.Value = listOrigin
        Else
            autoOriginCell.Value = vbNullString
        End If
        ApplyHiddenFormat design, autoOriginCell
    End If
End Sub

Private Sub ApplyHiddenFormat(ByVal design As ILLFormat, ByVal target As Range)
    If design Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat target, LinelistHiddenCell
    On Error GoTo 0
End Sub

Private Sub ApplyGeoStyling(ByVal context As ILLVarContext, _
                            ByVal layout As ILLVarLayout)
    Dim controlType As String
    Dim variableName As String
    Dim design As ILLFormat
    Dim labelCell As Range
    Dim valueCell As Range

    controlType = LCase$(Trim$(context.ValueOf("control")))
    variableName = LCase$(Trim$(context.ValueOf("variable name")))

    If Not ShouldApplyGeo(controlType, variableName) Then Exit Sub

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    Set labelCell = layout.LabelCell
    Set valueCell = layout.ValueCell

    On Error Resume Next
        If Not labelCell Is Nothing Then
            design.ApplyFormat labelCell, HListGeoHeader
        End If
        If Not valueCell Is Nothing Then
            design.ApplyFormat valueCell, HListGeo
        End If
    On Error GoTo 0
End Sub

Private Function ShouldApplyGeo(ByVal controlType As String, _
                                ByVal variableName As String) As Boolean
    Dim loweredControl As String
    Dim loweredName As String

    loweredControl = LCase$(Trim$(controlType))
    loweredName = LCase$(Trim$(variableName))

    ShouldApplyGeo = (InStr(loweredControl, "geo") > 0) _
                     Or (loweredControl = "hf") _
                     Or (InStr(loweredName, "pcode_adm") > 0)
End Function

Private Sub ApplyEditableLabel(ByVal context As ILLVarContext, _
                               ByVal layout As ILLVarLayout)
    Dim editableLabel As String
    Dim controlType As String
    Dim labelCell As Range

    editableLabel = context.EditableLabelFlag()
    controlType = LCase$(Trim$(context.ValueOf("control")))
    Set labelCell = layout.LabelCell

    If labelCell Is Nothing Then Exit Sub

    On Error Resume Next
        If (editableLabel = "yes") Or (controlType = "custom") Then
            labelCell.Locked = False
        Else
            labelCell.Locked = True
        End If
    On Error GoTo 0
End Sub

Private Sub ApplyRegisterBookRules(ByVal context As ILLVarContext, _
                                   ByVal layout As ILLVarLayout)
    Dim registerBook As String
    Dim statusValue As String
    Dim valueCell As Range
    Dim printedValue As Range

    registerBook = context.RegisterBookDirective()
    statusValue = LCase$(Trim$(context.ValueOf("status")))

    Set valueCell = layout.ValueCell
    Set printedValue = layout.PrintedValueCell

    ApplyLinelistColumnVisibility valueCell, statusValue
    ApplyPrintedColumnVisibility printedValue, registerBook, statusValue
    ApplyPrintedRotation printedValue, registerBook

    If Not valueCell Is Nothing Then
        On Error Resume Next
            valueCell.Locked = False
        On Error GoTo 0
    End If
End Sub

Private Sub ApplyLinelistColumnVisibility(ByVal valueCell As Range, _
                                          ByVal statusValue As String)
    If valueCell Is Nothing Then Exit Sub

    If (statusValue = "hidden") Or (statusValue = "optional, hidden") Then
        If valueCell.Column > 1 Then
            On Error Resume Next
                valueCell.EntireColumn.Hidden = True
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub ApplyPrintedColumnVisibility(ByVal printedValue As Range, _
                                         ByVal registerBook As String, _
                                         ByVal statusValue As String)
    If printedValue Is Nothing Then Exit Sub

    If (registerBook = "hidden") Or (statusValue = "hidden") Then
        On Error Resume Next
            printedValue.EntireColumn.Hidden = True
        On Error GoTo 0
    End If
End Sub

Private Sub ApplyPrintedRotation(ByVal printedValue As Range, _
                                 ByVal registerBook As String)
    Dim headerCell As Range

    If printedValue Is Nothing Then Exit Sub
    If registerBook <> "print, vertical header" Then Exit Sub

    Set headerCell = printedValue.Offset(-2, 0)

    On Error Resume Next
        If Not headerCell Is Nothing Then
            headerCell.Orientation = 90
            headerCell.EntireColumn.AutoFit
        End If
    On Error GoTo 0
End Sub

Private Function NormaliseControlType(ByVal rawControl As String) As String
    '@description Lowercase the control type and strip optional suffixes such as choice_multiple indices.
    Dim cleaned As String
    Dim bracketPos As Long

    cleaned = LCase$(Trim$(rawControl))
    bracketPos = InStr(cleaned, "(")
    If bracketPos > 0 Then
        cleaned = Left$(cleaned, bracketPos - 1)
    End If

    NormaliseControlType = cleaned
End Function

Private Sub HandleStandardChoices(ByVal context As ILLVarContext, _
                                  ByVal layout As ILLVarLayout, _
                                  ByVal controlType As String)
    '@description Materialise manual or multiple choice dropdowns with workbook-backed lists.
    Dim drop As IDropdownLists
    Dim categories As BetterArray
    Dim valueCell As Range
    Dim varName As String
    Dim alertType As String
    Dim messageText As String
    Dim showError As Boolean

    Set drop = ResolveDropdown(context, 1)
    If drop Is Nothing Then Exit Sub

    Set categories = ResolveCategories(context)
    If categories Is Nothing Then Exit Sub

    varName = context.VariableName
    drop.Add categories, varName

    Set valueCell = layout.ValueCell
    If valueCell Is Nothing Then Exit Sub

    alertType = context.ValueOf("alert")
    messageText = context.ValueOf("message")
    showError = (controlType <> "choice_multiple")

    ApplyDropdownValidation drop, valueCell, varName, alertType, messageText, showError
End Sub

Private Sub HandleCustomChoices(ByVal context As ILLVarContext, _
                                ByVal layout As ILLVarLayout)
    '@description Wire custom choice dropdowns, including hyperlinks and label restoration.
    Dim drop As IDropdownLists
    Dim categories As BetterArray
    Dim valueCell As Range
    Dim labelCell As Range
    Dim translation As ITranslationObject
    Dim customPrefix As String
    Dim varName As String
    Dim labelRange As Range
    Dim targetAddress As String
    Dim mainLabel As String
    Dim subLabel As String
    Dim mainLabelLength As Long
    Dim alertType As String
    Dim messageText As String
    Dim design As ILLFormat

    Set drop = ResolveDropdown(context, 2)
    If drop Is Nothing Then Exit Sub

    Set categories = ResolveCategories(context)
    If categories Is Nothing Then Exit Sub
    EnsureCategoryPresence categories

    Set valueCell = layout.ValueCell
    Set labelCell = layout.LabelCell
    If valueCell Is Nothing Then Exit Sub
    If labelCell Is Nothing Then Exit Sub

    Set translation = SafeTranslation(context)
    customPrefix = ResolveCustomPrefix(translation)

    varName = context.VariableName
    drop.Add categories, varName, addLabel:=True, counterPrefix:=customPrefix

    Set labelRange = ResolveLabelRangeSafe(drop, varName)
    If Not labelRange Is Nothing Then
        targetAddress = "'" & labelRange.Parent.Name & "'!" & labelRange.Address(False, False)

        On Error Resume Next
            labelCell.Parent.Hyperlinks.Add Anchor:=labelCell, Address:="", _
                                            SubAddress:=targetAddress, TextToDisplay:=""
            labelCell.Font.Underline = xlUnderlineStyleNone
        On Error GoTo 0

        mainLabel = context.ValueOf("main label")
        subLabel = context.ValueOf("sub label")
        mainLabelLength = Len(mainLabel)

        If LenB(subLabel) > 0 Then
            subLabel = subLabel & vbLf & "(" & CStr(labelRange.Value) & ")"
        Else
            subLabel = "(" & CStr(labelRange.Value) & ")"
        End If

        If LenB(mainLabel) > 0 Then
            labelCell.Value = mainLabel & vbLf & subLabel
        Else
            labelCell.Value = subLabel
        End If

        Set design = context.Design
        If Not design Is Nothing Then
            ApplyLabelScopes design, labelCell, mainLabelLength
        End If

        drop.AddReturnLink varName, labelCell
    End If

    alertType = context.ValueOf("alert")
    messageText = context.ValueOf("message")
    ApplyDropdownValidation drop, valueCell, varName, alertType, messageText, True
End Sub

Private Sub HandleListAuto(ByVal context As ILLVarContext, _
                           ByVal layout As ILLVarLayout)
    '@description Register list_auto dropdowns pointing to workbook-managed dynamic lists.
    Dim drop As IDropdownLists
    Dim categories As BetterArray
    Dim valueCell As Range
    Dim listName As String
    Dim alertType As String
    Dim messageText As String

    listName = Trim$(context.ValueOf("control details"))
    If LenB(listName) = 0 Then Exit Sub

    Set drop = ResolveDropdown(context, 1)
    If drop Is Nothing Then Exit Sub

    Set categories = New BetterArray
    categories.Push vbNullString
    drop.Add categories, listName

    Set valueCell = layout.ValueCell
    If valueCell Is Nothing Then Exit Sub

    alertType = context.ValueOf("alert")
    messageText = context.ValueOf("message")

    ApplyDropdownValidation drop, valueCell, listName, alertType, messageText, True
End Sub

Private Sub HandleGeoChoices(ByVal context As ILLVarContext, _
                             ByVal layout As ILLVarLayout)
    '@description Reproduce the chained geo dropdown workflow for admin levels 1-4.
    Dim drop As IDropdownLists
    Dim geo As ILLGeo
    Dim categories As BetterArray
    Dim valueCell As Range
    Dim alertType As String
    Dim messageText As String

    Set drop = ResolveDropdown(context, 1)
    If drop Is Nothing Then Exit Sub

    Set geo = ResolveGeo(context)
    If geo Is Nothing Then Exit Sub

    Set valueCell = layout.ValueCell
    If valueCell Is Nothing Then Exit Sub

    alertType = context.ValueOf("alert")
    messageText = context.ValueOf("message")

    Set categories = ResolveGeoLevel(geo, LevelAdmin1)
    If categories Is Nothing Then Exit Sub
    drop.Add categories, "admin1"
    ApplyDropdownValidation drop, valueCell, "admin1", alertType, messageText, True

    Set categories = New BetterArray
    categories.Push vbNullString

    drop.Add categories, "admin2"
    ApplyDropdownValidation drop, valueCell.Offset(0, 1), "admin2", alertType, messageText, True

    drop.Add categories, "admin3"
    ApplyDropdownValidation drop, valueCell.Offset(0, 2), "admin3", alertType, messageText, True

    drop.Add categories, "admin4"
    ApplyDropdownValidation drop, valueCell.Offset(0, 3), "admin4", alertType, messageText, True
End Sub

Private Function ResolveDropdown(ByVal context As ILLVarContext, _
                                 ByVal scope As Byte) As IDropdownLists
    On Error Resume Next
        Set ResolveDropdown = context.Dropdown(scope)
    On Error GoTo 0
End Function

Private Function ResolveCategories(ByVal context As ILLVarContext) As BetterArray
    On Error Resume Next
        Set ResolveCategories = context.Categories(context.VariableName)
    On Error GoTo 0
End Function

Private Sub EnsureCategoryPresence(ByVal categories As BetterArray)
    If categories Is Nothing Then Exit Sub
    If categories.Length = 0 Then
        categories.Push vbNullString
    End If
End Sub

Private Function ResolveCustomPrefix(ByVal translation As ITranslationObject) As String
    Dim resolved As String

    If translation Is Nothing Then
        resolved = "Custom choice"
    Else
        resolved = translation.TranslatedValue("MSG_CustomChoice")
        If LenB(resolved) = 0 Then
            resolved = "Custom choice"
        End If
    End If

    ResolveCustomPrefix = resolved
End Function

Private Function ResolveLabelRangeSafe(ByVal drop As IDropdownLists, _
                                       ByVal listName As String) As Range
    On Error Resume Next
        Set ResolveLabelRangeSafe = drop.LabelRange(listName)
    On Error GoTo 0
End Function

Private Sub ApplyDropdownValidation(ByVal drop As IDropdownLists, _
                                    ByVal target As Range, _
                                    ByVal listName As String, _
                                    ByVal alertType As String, _
                                    ByVal messageText As String, _
                                    ByVal showError As Boolean)
    If drop Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        drop.SetValidation cellRng:=target, listName:=listName, _
                           alertType:=alertType, message:=messageText, _
                           showError:=showError
    On Error GoTo 0
End Sub

Private Function IsCalculatedControl(ByVal controlType As String) As Boolean
    Select Case controlType
    Case "formula", "case_when", "choice_formula"
        IsCalculatedControl = True
    End Select
End Function

Private Function SafeDictionary(ByVal context As ILLVarContext) As ILLdictionary
    On Error Resume Next
        Set SafeDictionary = context.Dictionary
    On Error GoTo 0
End Function

Private Function SafeSpecs(ByVal context As ILLVarContext) As ILinelistSpecs
    On Error Resume Next
        Set SafeSpecs = context.Specs
    On Error GoTo 0
End Function

Private Function SafeTranslation(ByVal context As ILLVarContext) As ITranslationObject
    On Error Resume Next
        Set SafeTranslation = context.Translation
    On Error GoTo 0
End Function

Private Function ApplyParsedFormula(ByVal target As Range, _
                                    ByVal parsedFormula As String, _
                                    ByVal useArrayFormula As Boolean) As Boolean
    If target Is Nothing Then Exit Function
    If LenB(parsedFormula) = 0 Then Exit Function

    On Error GoTo PrimaryError
        If useArrayFormula Then
            target.FormulaArray = "=" & parsedFormula
        Else
            target.Formula = "=" & parsedFormula
        End If
        ApplyParsedFormula = True
        Exit Function

PrimaryError:
    Err.Clear
    If useArrayFormula Then
        On Error GoTo SecondaryError
            target.Formula = "=" & parsedFormula
            ApplyParsedFormula = True
            Exit Function
SecondaryError:
        Err.Clear
    End If
    ApplyParsedFormula = False
End Function

Private Sub ApplyFormulaFailure(ByVal target As Range)
    If target Is Nothing Then Exit Sub

    target.Value = Chr$(34) & "##== formula parsing failed :( ==##" & Chr$(34)
    target.Locked = True
End Sub

Private Sub ApplyCalculatedCellFormat(ByVal design As ILLFormat, _
                                      ByVal valueCell As Range)
    If design Is Nothing Then Exit Sub
    If valueCell Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat valueCell, HListCalculatedFormulaCell
    On Error GoTo 0
End Sub

Private Sub ApplyCalculatedHeaderFormat(ByVal design As ILLFormat, _
                                        ByVal labelCell As Range, _
                                        ByVal printedValueCell As Range)
    Dim printedHeader As Range

    If design Is Nothing Then Exit Sub

    On Error Resume Next
        If Not labelCell Is Nothing Then
            design.ApplyFormat labelCell, HListCalculatedFormulaHeader
        End If

        If Not printedValueCell Is Nothing Then
            Set printedHeader = printedValueCell.Offset(-2, 0)
            If Not printedHeader Is Nothing Then
                design.ApplyFormat printedHeader, HListCalculatedFormulaHeader
            End If
        End If
    On Error GoTo 0
End Sub

Private Sub ApplyFormattingCondition(ByVal context As ILLVarContext, _
                                     ByVal valueCell As Range)
    '@description Recreate conditional formatting driven by a supporting boolean column.
    Dim conditionName As String
    Dim vars As ILLVariables
    Dim templateCell As Range
    Dim conditionRow As Long
    Dim targetSheet As Worksheet
    Dim conditionSheet As Worksheet
    Dim conditionCell As Range
    Dim conditionFormula As String
    Dim conditions As FormatConditions
    Dim newCondition As FormatCondition

    If valueCell Is Nothing Then Exit Sub

    conditionName = context.FormattingConditionName()
    If LenB(conditionName) = 0 Then Exit Sub

    Set vars = SafeVariables(context)
    If vars Is Nothing Then Exit Sub
    If Not vars.Contains(conditionName, False) Then Exit Sub

    Set templateCell = ResolveFormattingTemplate(vars, context.VariableName)
    If templateCell Is Nothing Then Exit Sub

    conditionRow = ResolveVariableRow(vars, conditionName)
    If conditionRow <= 0 Then Exit Sub

    Set targetSheet = valueCell.Worksheet
    Set conditionSheet = ResolveConditionSheet(context, SafeSheetName(vars, conditionName), targetSheet)
    Set conditionCell = SafeCell(conditionSheet, conditionRow, valueCell.Column)
    If conditionCell Is Nothing Then Exit Sub

    conditionFormula = BuildConditionFormula(conditionCell, targetSheet)
    If LenB(conditionFormula) = 0 Then Exit Sub

    Set conditions = valueCell.FormatConditions
    RemoveExistingCondition conditions, conditionFormula

    On Error Resume Next
        conditions.Add Type:=xlExpression, Formula1:="=(" & conditionFormula & "= 1)"
    On Error GoTo 0

    If conditions Is Nothing Then Exit Sub
    If conditions.Count = 0 Then Exit Sub

    Set newCondition = conditions(conditions.Count)
    ApplyConditionTemplate newCondition, templateCell
End Sub

Private Sub ApplyUniqueCondition(ByVal context As ILLVarContext, _
                                 ByVal valueCell As Range)
    '@description Highlight duplicate entries when `unique = yes`.
    Dim uniqueFlag As String
    Dim conditions As FormatConditions
    Dim newCondition As FormatCondition

    If valueCell Is Nothing Then Exit Sub

    uniqueFlag = LCase$(Trim$(context.ValueOf("unique")))
    If uniqueFlag <> "yes" Then Exit Sub

    Set conditions = valueCell.FormatConditions

    On Error Resume Next
        conditions.AddUniqueValues
    On Error GoTo 0

    If conditions Is Nothing Then Exit Sub
    If conditions.Count = 0 Then Exit Sub

    Set newCondition = conditions(conditions.Count)

    On Error Resume Next
        newCondition.DupeUnique = xlDuplicate
        newCondition.Interior.Color = vbRed
    On Error GoTo 0
End Sub

Private Sub ApplyGeoWarning(ByVal context As ILLVarContext, _
                            ByVal valueCell As Range)
    '@description Surface geo mismatches by flagging non-empty entries where the pcode companion is blank.
    Dim controlType As String
    Dim warningName As String
    Dim vars As ILLVariables
    Dim warningRow As Long
    Dim warningSheet As Worksheet
    Dim warningCell As Range
    Dim actualFormula As String
    Dim warningFormula As String
    Dim combinedFormula As String
    Dim conditions As FormatConditions
    Dim newCondition As FormatCondition

    If valueCell Is Nothing Then Exit Sub

    controlType = LCase$(Trim$(context.ValueOf("control")))
    If InStr(controlType, "geo") = 0 Then Exit Sub

    warningName = "pcode_" & context.VariableName

    Set vars = SafeVariables(context)
    If vars Is Nothing Then Exit Sub
    If Not vars.Contains(warningName, False) Then Exit Sub

    warningRow = ResolveVariableRow(vars, warningName)
    If warningRow <= 0 Then Exit Sub

    Set warningSheet = ResolveConditionSheet(context, SafeSheetName(vars, warningName), valueCell.Worksheet)
    Set warningCell = SafeCell(warningSheet, warningRow, valueCell.Column)
    If warningCell Is Nothing Then Exit Sub

    warningFormula = BuildConditionFormula(warningCell, valueCell.Worksheet)
    actualFormula = BuildConditionFormula(valueCell, valueCell.Worksheet)

    If LenB(warningFormula) = 0 Then Exit Sub
    If LenB(actualFormula) = 0 Then Exit Sub

    combinedFormula = "=AND(" & warningFormula & "=" & Chr$(34) & Chr$(34) & "," & _
                      actualFormula & "<>" & Chr$(34) & Chr$(34) & ")"

    Set conditions = valueCell.FormatConditions

    On Error Resume Next
        conditions.Add Type:=xlExpression, Formula1:=combinedFormula
    On Error GoTo 0

    If conditions Is Nothing Then Exit Sub
    If conditions.Count = 0 Then Exit Sub

    Set newCondition = conditions(conditions.Count)

    On Error Resume Next
        newCondition.Interior.Color = RGB(242, 147, 12)
        newCondition.SetFirstPriority
        newCondition.StopIfTrue = True
    On Error GoTo 0
End Sub

Private Function ResolveFormattingTemplate(ByVal vars As ILLVariables, _
                                           ByVal varName As String) As Range
    On Error Resume Next
        Set ResolveFormattingTemplate = vars.CellRange("formatting values", varName)
    On Error GoTo 0
End Function

Private Function ResolveVariableRow(ByVal vars As ILLVariables, _
                                    ByVal varName As String) As Long
    On Error Resume Next
        ResolveVariableRow = vars.Index(varName)
    On Error GoTo 0
End Function

Private Function ResolveConditionSheet(ByVal context As ILLVarContext, _
                                       ByVal sheetName As String, _
                                       ByVal fallbackSheet As Worksheet) As Worksheet
    Dim linelist As ILinelist
    Dim targetSheet As Worksheet

    If LenB(sheetName) = 0 Then
        Set ResolveConditionSheet = fallbackSheet
        Exit Function
    End If

    If Not fallbackSheet Is Nothing Then
        If StrComp(fallbackSheet.Name, sheetName, vbTextCompare) = 0 Then
            Set ResolveConditionSheet = fallbackSheet
            Exit Function
        End If
    End If

    Set linelist = SafeLinelist(context)
    If Not linelist Is Nothing Then
        On Error Resume Next
            Set targetSheet = linelist.Wksh(sheetName)
        On Error GoTo 0
    End If

    If targetSheet Is Nothing Then
        On Error Resume Next
            Set targetSheet = fallbackSheet.Parent.Worksheets(sheetName)
        On Error GoTo 0
    End If

    If targetSheet Is Nothing Then
        Set targetSheet = fallbackSheet
    End If

    Set ResolveConditionSheet = targetSheet
End Function

Private Function SafeCell(ByVal sheet As Worksheet, ByVal rowIndex As Long, ByVal columnIndex As Long) As Range
    If sheet Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function
    If columnIndex <= 0 Then Exit Function

    On Error Resume Next
        Set SafeCell = sheet.Cells(rowIndex, columnIndex)
    On Error GoTo 0
End Function

Private Function BuildConditionFormula(ByVal conditionCell As Range, _
                                       ByVal anchorSheet As Worksheet) As String
    Dim sheetPrefix As String
    Dim targetSheet As Worksheet
    Dim addressText As String

    If conditionCell Is Nothing Then Exit Function

    Set targetSheet = conditionCell.Worksheet
    If targetSheet Is Nothing Then Exit Function

    addressText = conditionCell.Address(RowAbsolute:=True, ColumnAbsolute:=True)

    sheetPrefix = "'" & EscapeSheetName(targetSheet.Name) & "'!"

    BuildConditionFormula = sheetPrefix & addressText
End Function

Private Function EscapeSheetName(ByVal sheetName As String) As String
    EscapeSheetName = Replace(sheetName, "'", "''")
End Function

Private Sub RemoveExistingCondition(ByVal conditions As FormatConditions, _
                                    ByVal conditionFormula As String)
    Dim idx As Long
    Dim candidate As String
    Dim targetSignature As String

    If conditions Is Nothing Then Exit Sub
    If conditions.Count = 0 Then Exit Sub

    targetSignature = NormaliseConditionFormula("=(" & conditionFormula & "=1)")

    For idx = conditions.Count To 1 Step -1
        On Error Resume Next
            candidate = conditions(idx).Formula1
        On Error GoTo 0

        If LenB(candidate) > 0 Then
            If StrComp(NormaliseConditionFormula(candidate), targetSignature, vbTextCompare) = 0 Then
                On Error Resume Next
                    conditions(idx).Delete
                On Error GoTo 0
            End If
        End If
    Next idx
End Sub

Private Function NormaliseConditionFormula(ByVal formulaText As String) As String
    Dim working As String
    working = Replace(Trim$(formulaText), " ", vbNullString)
    NormaliseConditionFormula = LCase$(working)
End Function

Private Sub ApplyConditionTemplate(ByVal condition As FormatCondition, _
                                   ByVal template As Range)
    If condition Is Nothing Then Exit Sub
    If template Is Nothing Then Exit Sub

    On Error Resume Next
        condition.Interior.Pattern = template.Interior.Pattern
        condition.Interior.Color = template.Interior.Color
        condition.Font.Color = template.Font.Color
        condition.Font.Bold = template.Font.Bold
        condition.Font.Italic = template.Font.Italic
        condition.SetFirstPriority
        condition.StopIfTrue = True
    On Error GoTo 0
End Sub

Private Function SafeVariables(ByVal context As ILLVarContext) As ILLVariables
    On Error Resume Next
        Set SafeVariables = context.Variables
    On Error GoTo 0
End Function

Private Function SafeLinelist(ByVal context As ILLVarContext) As ILinelist
    On Error Resume Next
        Set SafeLinelist = context.Linelist
    On Error GoTo 0
End Function

Private Function SafeSheetName(ByVal vars As ILLVariables, _
                               ByVal varName As String) As String
    On Error Resume Next
        SafeSheetName = vars.Value("sheet name", varName)
    On Error GoTo 0
End Function

Private Function ResolveGeo(ByVal context As ILLVarContext) As ILLGeo
    Dim specs As ILinelistSpecs

    Set specs = SafeSpecs(context)
    If specs Is Nothing Then Exit Function

    On Error Resume Next
        Set ResolveGeo = specs.GeoObject
    On Error GoTo 0
End Function

Private Function ResolveGeoLevel(ByVal geo As ILLGeo, ByVal adminLevel As Long) As BetterArray
    On Error Resume Next
        Set ResolveGeoLevel = geo.GeoLevel(adminLevel)
    On Error GoTo 0
End Function
