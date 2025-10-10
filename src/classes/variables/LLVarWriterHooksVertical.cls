VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLVarWriterHooksVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Default hooks for vertical variable writer")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit
Implements ILLVarWriterHooks

Private Const START_ANCHOR_INDEX As Long = 4

Private Sub ILLVarWriterHooks_AfterLabel(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout, _
                                         ByVal labelCell As Range, _
                                         ByVal mainLabelLength As Long)
    Dim design As ILLFormat
    Dim controlType As String

    controlType = LCase$(Trim$(context.ValueOf("control")))
    Set design = context.Design

    If Not design Is Nothing Then
        If Not labelCell Is Nothing Then
            If mainLabelLength > 0 Then
                If controlType <> "choice_custom" Then
                    ApplyLabelFormatting design, labelCell, mainLabelLength
                End If
            End If
        End If
    End If

    ApplyStartAnchor context, layout
End Sub

Private Sub ILLVarWriterHooks_AfterType(ByVal context As ILLVarContext, _
                                        ByVal layout As ILLVarLayout, _
                                        ByVal valueCell As Range)
    'No additional behaviour required
End Sub

Private Sub ILLVarWriterHooks_AfterValidation(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout, _
                                              ByVal valueCell As Range)
    ApplyFormattingCondition context, layout, valueCell
End Sub

Private Sub ILLVarWriterHooks_AfterWrite(ByVal context As ILLVarContext, _
                                         ByVal layout As ILLVarLayout)
    RepairLabelAfterChoices context, layout
    PopulateHiddenMetadata context, layout
    ApplyRowVisibility context, layout
End Sub

Private Sub ILLVarWriterHooks_ConfigureLayout(ByVal context As ILLVarContext, _
                                              ByVal layout As ILLVarLayout)
    Dim anchorRow As Long
    Dim sheet As Worksheet

    anchorRow = CLng(context.ValueOf("column index"))
    Set sheet = context.Worksheet

    layout.Configure sheet, anchorRow
End Sub

Private Sub ILLVarWriterHooks_WriteChoices(ByVal context As ILLVarContext, _
                                           ByVal layout As ILLVarLayout)
    Dim controlType As String

    controlType = LCase$(Trim$(context.ValueOf("control")))

    Select Case controlType
    Case "choice_manual", "choice_multiple"
        HandleStandardChoices context, layout, controlType

    Case "choice_custom"
        HandleCustomChoices context, layout
    End Select
End Sub

Private Sub ILLVarWriterHooks_WriteFormulas(ByVal context As ILLVarContext, _
                                            ByVal layout As ILLVarLayout)
    Dim controlType As String
    Dim controlDetails As String
    Dim valueCell As Range
    Dim dict As ILLdictionary
    Dim specs As ILinelistSpecs
    Dim formulaData As IFormulaData
    Dim formulaEngine As IFormulas
    Dim parsedFormula As String
    Dim isGroupedFormula As Boolean

    controlType = LCase$(Trim$(context.ValueOf("control")))
    If Not IsCalculatedControl(controlType) Then Exit Sub

    controlDetails = Trim$(context.ValueOf("control details"))
    If LenB(controlDetails) = 0 Then Exit Sub

    Set valueCell = layout.ValueCell
    If valueCell Is Nothing Then Exit Sub

    Set dict = SafeDictionary(context)
    Set specs = SafeSpecs(context)
    If specs Is Nothing Then Exit Sub

    On Error Resume Next
        Set formulaData = specs.FormulaDataObject
    On Error GoTo 0

    If formulaData Is Nothing Then Exit Sub
    If dict Is Nothing Then Exit Sub

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
    ApplyCalculatedFormat context.Design, valueCell
End Sub

Private Sub ILLVarWriterHooks_ConfigureProtection(ByVal context As ILLVarContext, _
                                                  ByVal layout As ILLVarLayout, _
                                                  ByVal valueCell As Range, _
                                                  ByVal metadataRange As Range)
    Dim editableLabel As String
    Dim labelCell As Range

    editableLabel = context.EditableLabelFlag()
    Set labelCell = layout.LabelCell

    If Not valueCell Is Nothing Then
        If editableLabel = "yes" Then
            LockCell valueCell.Offset(0, -1), False
        Else
            LockCell valueCell.Offset(0, -1), True
        End If

        LockCell valueCell.Offset(0, -2), True
        LockCell valueCell.Offset(0, -3), True
    End If

    If Not labelCell Is Nothing Then
        LockCell labelCell, Not (editableLabel = "yes")
    End If
End Sub

Private Sub ApplyLabelFormatting(ByVal design As ILLFormat, _
                                 ByVal labelCell As Range, _
                                 ByVal mainLabelLength As Long)

    Dim totalLength As Long

    On Error Resume Next
        design.ApplyFormat labelCell.Characters(Start:=1, Length:=mainLabelLength), _
                           scope:=VListMainLab
        totalLength = Len(CStr(labelCell.Value))
        If totalLength > mainLabelLength Then
            design.ApplyFormat labelCell.Characters(Start:=mainLabelLength + 1), _
                               scope:=VListSublab
        End If
    On Error GoTo 0
End Sub

Private Sub ApplyStartAnchor(ByVal context As ILLVarContext, _
                             ByVal layout As ILLVarLayout)

    Dim columnIndex As Long
    Dim tableName As String
    Dim anchorCell As Range

    columnIndex = ResolveColumnIndex(context.ValueOf("column index"))
    If columnIndex <> START_ANCHOR_INDEX Then Exit Sub

    tableName = Trim$(context.ValueOf("table name"))
    If LenB(tableName) = 0 Then Exit Sub

    Set anchorCell = layout.AnchorCell
    AssignWorkbookName anchorCell, tableName & "_START"
End Sub

Private Function ResolveColumnIndex(ByVal indexValue As String) As Long
    Dim resolved As Long

    On Error Resume Next
        resolved = CLng(Trim$(indexValue))
    On Error GoTo 0

    ResolveColumnIndex = resolved
End Function

Private Sub AssignWorkbookName(ByVal target As Range, ByVal rangeName As String)
    Dim wb As Workbook

    If target Is Nothing Then Exit Sub
    If LenB(rangeName) = 0 Then Exit Sub

    Set wb = target.Worksheet.Parent

    On Error Resume Next
        wb.Names(rangeName).Delete
    On Error GoTo 0

    On Error Resume Next
        wb.Names.Add Name:=rangeName, RefersTo:=target
    On Error GoTo 0
End Sub

Private Sub LockCell(ByVal target As Range, ByVal shouldLock As Boolean)
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        target.Locked = shouldLock
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

Private Sub ApplyCalculatedFormat(ByVal design As ILLFormat, ByVal target As Range)
    If design Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat target, HListCalculatedFormulaCell
    On Error GoTo 0
End Sub

Private Sub ApplyFormattingCondition(ByVal context As ILLVarContext, _
                                     ByVal layout As ILLVarLayout, _
                                     ByVal valueCell As Range)
    '@description Replicate legacy conditional formatting that highlights rows when the formatting condition evaluates to 1.
    Dim conditionName As String
    Dim vars As ILLVariables
    Dim templateCell As Range
    Dim targetSheet As Worksheet
    Dim conditionSheet As Worksheet
    Dim conditionRow As Long
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
    Set conditionSheet = ResolveConditionSheet(context, vars.Value("sheet name", conditionName), targetSheet)
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

Private Sub RepairLabelAfterChoices(ByVal context As ILLVarContext, ByVal layout As ILLVarLayout)
    '@description Reapply label styling for custom choice hyperlinks that reset formatting.
    Dim controlType As String
    Dim design As ILLFormat
    Dim labelCell As Range
    Dim mainLabel As String

    controlType = LCase$(Trim$(context.ValueOf("control")))
    If controlType <> "choice_custom" Then Exit Sub

    Set design = context.Design
    If design Is Nothing Then Exit Sub

    Set labelCell = layout.LabelCell
    If labelCell Is Nothing Then Exit Sub

    mainLabel = context.ValueOf("main label")
    If LenB(mainLabel) = 0 Then Exit Sub

    ApplyLabelFormatting design, labelCell, Len(mainLabel)
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

    If anchorSheet Is Nothing Then
        sheetPrefix = "'" & EscapeSheetName(targetSheet.Name) & "'!"
    ElseIf Not (targetSheet Is anchorSheet) Then
        sheetPrefix = "'" & EscapeSheetName(targetSheet.Name) & "'!"
    Else
        sheetPrefix = "'" & EscapeSheetName(targetSheet.Name) & "'!"
    End If

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

Private Sub HandleStandardChoices(ByVal context As ILLVarContext, _
                                  ByVal layout As ILLVarLayout, _
                                  ByVal controlType As String)

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

    Dim drop As IDropdownLists
    Dim categories As BetterArray
    Dim translation As ITranslationObject
    Dim valueCell As Range
    Dim labelCell As Range
    Dim varName As String
    Dim customLabel As String
    Dim labelRange As Range
    Dim targetAddress As String
    Dim mainLabel As String
    Dim subLabel As String
    Dim mainLabelLength As Long
    Dim design As ILLFormat
    Dim alertType As String
    Dim messageText As String

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
    customLabel = ResolveCustomPrefix(translation)

    varName = context.VariableName
    drop.Add categories, varName, addLabel:=True, counterPrefix:=customLabel

    Set labelRange = ResolveLabelRangeSafe(drop, varName)
    If Not labelRange Is Nothing Then
        targetAddress = "'" & labelRange.Parent.Name & "'!" & labelRange.Address(False, False)

        On Error Resume Next
            labelCell.Parent.Hyperlinks.Add Anchor:=labelCell, Address:="", _
                                            SubAddress:=targetAddress, TextToDisplay:=""
        On Error GoTo 0

        On Error Resume Next
            labelCell.Font.Underline = xlUnderlineStyleNone
        On Error GoTo 0

        mainLabel = context.ValueOf("main label")
        subLabel = context.ValueOf("sub label")
        mainLabelLength = Len(mainLabel)
        subLabel = subLabel & vbLf & "(" & CStr(labelRange.Value) & ")"

        labelCell.Value = mainLabel & vbLf & subLabel

        Set design = context.Design
        If Not design Is Nothing Then
            ApplyLabelFormatting design, labelCell, mainLabelLength
        End If

        drop.AddReturnLink varName, labelCell
    End If

    alertType = context.ValueOf("alert")
    messageText = context.ValueOf("message")
    ApplyDropdownValidation drop, valueCell, varName, alertType, messageText, True
End Sub

Private Function ResolveCategories(ByVal context As ILLVarContext) As BetterArray
    Dim categories As BetterArray

    On Error Resume Next
        Set categories = context.Categories(context.VariableName)
    On Error GoTo 0

    Set ResolveCategories = categories
End Function

Private Function ResolveDropdown(ByVal context As ILLVarContext, _
                                 ByVal scope As Byte) As IDropdownLists
    On Error Resume Next
        Set ResolveDropdown = context.Dropdown(scope)
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
        If LenB(resolved) = 0 Then resolved = "Custom choice"
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


Private Sub PopulateHiddenMetadata(ByVal context As ILLVarContext, _
                                   ByVal layout As ILLVarLayout)
    Dim controlValue As String
    Dim controlCell As Range
    Dim design As ILLFormat

    controlValue = context.ValueOf("control")
    Set controlCell = layout.ControlCell
    Set design = context.Design

    If Not controlCell Is Nothing Then
        controlCell.Value = controlValue
        ApplyHiddenDesign design, controlCell
    End If
End Sub

Private Sub ApplyHiddenDesign(ByVal design As ILLFormat, ByVal target As Range)
    If design Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub

    On Error Resume Next
        design.ApplyFormat target, LinelistHiddenCell
    On Error GoTo 0
End Sub

Private Sub ApplyRowVisibility(ByVal context As ILLVarContext, _
                               ByVal layout As ILLVarLayout)
    Dim statusValue As String
    Dim hiddenRange As Range

    statusValue = LCase$(Trim$(context.ValueOf("status")))

    If (statusValue <> "hidden") And (statusValue <> "optional, hidden") Then Exit Sub

    Set hiddenRange = layout.HiddenRange
    If hiddenRange Is Nothing Then Exit Sub

    On Error Resume Next
        hiddenRange.EntireRow.Hidden = True
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
