Attribute VB_Name = "TestLLVarWriterBase"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private lastSpecsStub As LLVarContextSpecsStub
Private lastVariablesStub As LLVarContextVariablesStub
Private lastLinelistStub As LLVarContextLinelistStub

Private Const WRITER_SHEET As String = "LLVarWriterSheet"
Private Const WRITER_PRINT_SHEET As String = "LLVarWriterSheet_Print"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet WRITER_SHEET
    TestHelpers.DeleteWorksheet WRITER_PRINT_SHEET
    Set Assert = Nothing
End Sub

Private Function BuildContext(ByVal varType As String, _
                              Optional ByVal varFormat As String = vbNullString, _
                              Optional ByVal minValue As String = vbNullString, _
                              Optional ByVal maxValue As String = vbNullString, _
                              Optional ByVal controlType As String = "text", _
                              Optional ByVal columnIndex As Long = 3, _
                              Optional ByVal tableName As String = "tbl_test", _
                              Optional ByVal editableLabel As String = "no", _
                              Optional ByVal design As ILLFormat = Nothing, _
                              Optional ByVal statusValue As String = "active", _
                              Optional ByVal controlDetails As String = vbNullString, _
                              Optional ByVal formattingCondition As String = vbNullString, _
                              Optional ByVal dropdownDefault As IDropdownLists = Nothing, _
                              Optional ByVal dropdownCustom As IDropdownLists = Nothing, _
                              Optional ByVal dictionary As ILLdictionary = Nothing, _
                              Optional ByVal formulaData As IFormulaData = Nothing, _
                              Optional ByVal translation As ITranslationObject = Nothing, _
                              Optional ByVal listAutoOrigin As String = vbNullString, _
                              Optional ByVal registerBook As String = vbNullString, _
                              Optional ByVal variableName As String = "test_var") As ILLVarContext

    Dim linelistStub As LLVarContextLinelistStub
    Dim variablesStub As LLVarContextVariablesStub
    Dim specsStub As LLVarContextSpecsStub
    Dim context As ILLVarContext

    Set linelistStub = New LLVarContextLinelistStub
    linelistStub.UseWorkbook ThisWorkbook

    Set specsStub = New LLVarContextSpecsStub
    linelistStub.UseSpecs specsStub
    If Not design Is Nothing Then
        specsStub.SetDesignFormat design
    End If
    If Not dictionary Is Nothing Then
        specsStub.SetDictionary dictionary
        linelistStub.UseDictionary dictionary
    End If
    If Not formulaData Is Nothing Then
        specsStub.SetFormulaData formulaData
    End If
    If Not translation Is Nothing Then
        specsStub.SetTranslation translation
    End If
    If (Not dropdownDefault Is Nothing) Or (Not dropdownCustom Is Nothing) Then
        linelistStub.UseDropdowns dropdownDefault, dropdownCustom
    End If

    Set variablesStub = New LLVarContextVariablesStub
    variablesStub.AddValue variableName, "sheet name", WRITER_SHEET
    variablesStub.AddValue variableName, "column index", CStr(columnIndex)
    variablesStub.AddValue variableName, "main label", "Main"
    variablesStub.AddValue variableName, "sub label", "Sub"
    variablesStub.AddValue variableName, "variable name", variableName
    variablesStub.AddValue variableName, "variable type", varType
    variablesStub.AddValue variableName, "variable format", varFormat
    variablesStub.AddValue variableName, "status", statusValue
    variablesStub.AddValue variableName, "note", "A helpful note"
    variablesStub.AddValue variableName, "control", controlType
    variablesStub.AddValue variableName, "min", minValue
    variablesStub.AddValue variableName, "max", maxValue
    variablesStub.AddValue variableName, "alert", "warning"
    variablesStub.AddValue variableName, "message", "Range check"
    variablesStub.AddValue variableName, "table name", tableName
    variablesStub.AddValue variableName, "editable label", editableLabel
    variablesStub.AddValue variableName, "list auto", listAutoOrigin
    variablesStub.AddValue variableName, "register book", registerBook
    If LenB(controlDetails) > 0 Then
        variablesStub.AddValue variableName, "control details", controlDetails
    End If
    If LenB(formattingCondition) > 0 Then
        variablesStub.AddValue variableName, "formatting condition", formattingCondition
    End If

    Set context = New LLVarContext
    context.Initialise linelistStub, "test_var", , variablesStub, specsStub

    Set lastSpecsStub = specsStub
    Set lastVariablesStub = variablesStub
    Set lastLinelistStub = linelistStub

    Set BuildContext = context
End Function

Private Function BuildWriter(ByVal context As ILLVarContext) As ILLVarWriter
    Dim layout As ILLVarLayout
    Dim hooks As ILLVarWriterHooks
    Dim writer As ILLVarWriter

    Set layout = New LLVarLayoutHorizontal
    Set hooks = New LLVarWriterHooksStub
    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks

    Set BuildWriter = writer
End Function

Private Function CurrentSpecsStub() As LLVarContextSpecsStub
    Set CurrentSpecsStub = lastSpecsStub
End Function

Private Function CurrentVariablesStub() As LLVarContextVariablesStub
    Set CurrentVariablesStub = lastVariablesStub
End Function

Private Function CurrentLinelistStub() As LLVarContextLinelistStub
    Set CurrentLinelistStub = lastLinelistStub
End Function

Private Function RangeNameOrEmpty(ByVal target As Range) As String
    On Error Resume Next
        RangeNameOrEmpty = target.Name.Name
    On Error GoTo 0
End Function

'@TestMethod("LLVarWriter")
Private Sub TestWriteVariablePopulatesLabels()
    Dim context As ILLVarContext
    Dim writer As ILLVarWriter
    Dim sheet As Worksheet
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text")
    Set hooks = New LLVarWriterHooksStub
    Set layout = New LLVarLayoutHorizontal

    layout.Configure sheet, 3

    Dim base As LLVarWriterBase
    Set base = New LLVarWriterBase
    base.Initialise context, layout, hooks
    base.WriteVariable

    Assert.AreEqual "Main" & vbLf & "Sub", layout.LabelCell.Value, _
                     "Writer should combine main and sub labels"
    Assert.AreEqual "test_var", layout.NameCell.Value, _
                     "Variable name should be written next to the label"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestWriteVariableAppliesNumberFormat()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksStub

    TestHelpers.EnsureWorksheet WRITER_SHEET

    Set context = BuildContext("decimal")
    Set layout = New LLVarLayoutHorizontal
    layout.Configure ThisWorkbook.Worksheets(WRITER_SHEET), 3
    Set hooks = New LLVarWriterHooksStub

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "#,##0.00;-#,##0.00;0.00;@", layout.ValueCell.NumberFormat, _
                     "Default decimal number format should be applied"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestWriteVariableAddsValidation()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksStub

    TestHelpers.EnsureWorksheet WRITER_SHEET

    Set context = BuildContext("integer", , "1", "10")
    Set layout = New LLVarLayoutHorizontal
    layout.Configure ThisWorkbook.Worksheets(WRITER_SHEET), 3
    Set hooks = New LLVarWriterHooksStub

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual xlValidateWholeNumber, layout.ValueCell.Validation.Type, _
                     "Whole number validation should be applied when range limits exist"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestWriteVariableUnlocksValueCell()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksStub
    Dim sheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("integer")
    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksStub

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.IsFalse layout.ValueCell.Locked, _
                   "Value cells should remain unlocked so users can enter linelist data"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestValidationConvertsFormulaToLiteral()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksStub
    Dim sheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("integer", vbNullString, "=1+1")
    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksStub

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "2", layout.ValueCell.Validation.Formula1, _
                     "Formula thresholds should be resolved to literal values when possible"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterNamesValueCell()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksStub
    Dim sheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text")
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksStub

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual vbNullString, layout.ValueCell.Value, _
                     "Vertical layout should keep the input cell empty after naming the range"
    Assert.AreEqual "test_var", RangeNameOrEmpty(layout.ValueCell), _
                     "Vertical layout should assign the variable name to the value cell"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestWriteVariableAppliesCurrencyFormat()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksStub
    Dim sheet As Worksheet
    Dim expectedFormat As String

    expectedFormat = "#,##0.00 [$€-x-euro1];-#,##0.00 [$€-x-euro1];0.00 [$€-x-euro1];@"

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("decimal", "euros")
    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksStub

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual expectedFormat, layout.ValueCell.NumberFormat, _
                     "Currency variables marked as euros should use the dedicated number format"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterAppliesLabelStyling()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet
    Dim design As LLFormatLogStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set design = New LLFormatLogStub
    Set context = BuildContext("text", design:=design)
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual 1, design.ScopeCount(VListMainLab), _
                     "Vertical hooks should format the main label using VListMainLab"
    Assert.AreEqual 1, design.ScopeCount(VListSublab), _
                     "Vertical hooks should format the sub label using VListSublab"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterCreatesStartAnchor()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet
    Dim startName As Name
    Dim expectedName As String

    expectedName = "tbl_anchor_START"

    On Error Resume Next
        ThisWorkbook.Names(expectedName).Delete
    On Error GoTo 0

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text", columnIndex:=4, tableName:="tbl_anchor")
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 4
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    On Error Resume Next
        Set startName = ThisWorkbook.Names(expectedName)
    On Error GoTo 0

    Assert.IsFalse startName Is Nothing, _
                   "Vertical hooks should register the table start anchor when column index equals 4"
    Assert.AreEqual layout.ValueCell.Address(False, False), _
                     startName.RefersToRange.Address(False, False), _
                     "Start anchor must point to the variable value cell"

    On Error Resume Next
        ThisWorkbook.Names(expectedName).Delete
    On Error GoTo 0
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterUnlocksEditableLabel()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text", editableLabel:="yes")
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.IsFalse layout.LabelCell.Locked, _
                   "Editable labels should unlock the label cell for user edits"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterPopulatesControlMetadata()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text", controlType:="choice_manual")
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "choice_manual", layout.ControlCell.Value, _
                     "Control metadata should be stored next to the value cell"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterHidesHiddenRows()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text", statusValue:="hidden")
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.IsTrue layout.HiddenRange.EntireRow.Hidden, _
                  "Rows flagged as hidden should be hidden on the worksheet"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterFormatsLabels()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim design As LLFormatLogStub
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim printedLabel As Range

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set design = New LLFormatLogStub
    Set context = BuildContext("text", design:=design)
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Set printedLabel = layout.PrintedValueCell.Offset(-2, 0)

    Assert.AreEqual layout.LabelCell.Value, printedLabel.Value, _
                     "Printed sheet should mirror the combined linelist header text"
    Assert.AreEqual context.ValueOf("variable name"), _
                     layout.PrintedValueCell.Offset(-1, 0).Value, _
                     "Printed sheet should expose the variable name above the entry cell"
    Assert.IsTrue design.ScopeCount(HListMainLab) >= 2, _
                  "Main label styling should be applied on both linelist and printed headers"
    Assert.IsTrue design.ScopeCount(HListSublab) >= 2, _
                  "Sub label styling should be applied on both linelist and printed headers"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterPopulatesMetadata()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim design As LLFormatLogStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set design = New LLFormatLogStub
    Set context = BuildContext("text", controlType:="choice_custom", design:=design, _
                               listAutoOrigin:="list origin")
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3, printSheet
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "choice_custom", layout.ControlCell.Value, _
                     "Horizontal hooks should record the control type next to the value cell"
    Assert.AreEqual "list origin", layout.AutoOriginCell.Value, _
                     "Horizontal hooks should capture the list auto origin metadata"
    Assert.IsTrue design.ScopeCount(LinelistHiddenCell) >= 2, _
                  "Hidden cells should use the LinelistHiddenCell format (control and list origin)"
    Assert.IsFalse layout.LabelCell.Locked, _
                   "Editable/custom controls should unlock the label header"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAppliesGeoStyling()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim design As LLFormatLogStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set design = New LLFormatLogStub
    Set context = BuildContext("text", controlType:="geo1", design:=design)
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3, printSheet
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.IsTrue design.ScopeCount(HListGeoHeader) >= 1, _
                  "Geo variables should apply geo header styling"
    Assert.IsTrue design.ScopeCount(HListGeo) >= 1, _
                  "Geo variables should apply geo data styling"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAppliesRegisterBookRules()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set context = BuildContext("text", columnIndex:=2, statusValue:="optional, hidden", _
                               registerBook:="print, vertical header")
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 2, printSheet
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.IsTrue layout.ValueCell.EntireColumn.Hidden, _
                  "Optional hidden columns should be hidden on the linelist sheet"
    Assert.AreEqual 90, layout.PrintedValueCell.Offset(-2, 0).Orientation, _
                       "Register book 'print, vertical header' should rotate printed headers"
    Assert.IsFalse layout.ValueCell.Locked, _
                  "Horizontal entry cells should remain unlocked for data entry"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterHidesPrintedWhenRequested()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set context = BuildContext("text", registerBook:="hidden")
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3, printSheet
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.IsTrue layout.PrintedValueCell.EntireColumn.Hidden, _
                  "Register book 'hidden' should hide the printed worksheet column"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterCreatesAnchors()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim startAnchor As Name
    Dim printAnchor As Name
    Const TABLE_NAME As String = "tbl_horizontal_anchor"

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    On Error Resume Next
        ThisWorkbook.Names(TABLE_NAME & "_START").Delete
        ThisWorkbook.Names(TABLE_NAME & "_PRINTSTART").Delete
    On Error GoTo 0

    Set context = BuildContext("text", columnIndex:=1, tableName:=TABLE_NAME)
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 1
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Set startAnchor = ThisWorkbook.Names(TABLE_NAME & "_START")
    Set printAnchor = ThisWorkbook.Names(TABLE_NAME & "_PRINTSTART")

    Assert.IsFalse startAnchor Is Nothing, _
                  "_START anchor should be stored on the linelist header"
    Assert.AreEqual layout.NameCell.Address(False, False), _
                     startAnchor.RefersToRange.Address(False, False), _
                     "_START anchor should target the linelist name cell"
    Assert.IsFalse printAnchor Is Nothing, _
                  "_PRINTSTART anchor should be stored on the printed header"
    Assert.AreEqual layout.PrintedValueCell.Offset(-1, 0).Address(False, False), _
                     printAnchor.RefersToRange.Address(False, False), _
                     "_PRINTSTART anchor should target the printed name cell"

    On Error Resume Next
        startAnchor.Delete
        printAnchor.Delete
    On Error GoTo 0
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsManualChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim categories As BetterArray

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set context = BuildContext("text", controlType:="choice_manual", dropdownDefault:=dropStub)

    Set categories = New BetterArray
    categories.LowerBound = 1
    categories.Push "Option A"
    categories.Push "Option B"
    CurrentSpecsStub.SetCategories "test_var", categories

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "test_var", dropStub.LastAddedListName, _
                     "Manual choices should register the dropdown by variable name"
    Assert.AreEqual 1, dropStub.ValidationCallCount, _
                     "Manual choices should apply a single validation to the value cell"
    Assert.AreEqual "test_var", dropStub.ValidationListName(1), _
                     "Manual choice validation should reference the variable list name"
    Assert.IsTrue dropStub.ValidationShowError(1), _
                  "Manual choice validation should be configured to display errors"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsMultipleChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim categories As BetterArray

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set context = BuildContext("text", controlType:="choice_multiple", dropdownDefault:=dropStub)

    Set categories = New BetterArray
    categories.LowerBound = 1
    categories.Push "Code A"
    categories.Push "Code B"
    CurrentSpecsStub.SetCategories "test_var", categories

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "test_var", dropStub.ValidationListName(1), _
                     "Multiple choice validation should still target the variable list"
    Assert.IsFalse dropStub.ValidationShowError(1), _
                   "Multiple choice validation should suppress in-cell error prompts"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsCustomChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim categories As BetterArray
    Dim translation As LinelistSpecsTranslationStub
    Dim design As LLFormatLogStub
    Dim labelText As String

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set translation = New LinelistSpecsTranslationStub
    translation.Initialise
    translation.SetTranslation "MSG_CustomChoice", "Custom"

    Set design = New LLFormatLogStub

    Set context = BuildContext("text", controlType:="choice_custom", dropdownDefault:=dropStub, _
                               dropdownCustom:=dropStub, design:=design, translation:=translation)
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set categories = New BetterArray
    categories.LowerBound = 1
    categories.Push "Custom Option"
    CurrentSpecsStub.SetCategories "test_var", categories

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3, printSheet
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    labelText = CStr(layout.LabelCell.Value)

    Assert.IsTrue InStr(labelText, "(") > 0, _
                  "Custom dropdowns should append the generated entry label to the sub label"
    Assert.AreEqual 1, layout.LabelCell.Hyperlinks.Count, _
                     "Custom dropdowns should create a hyperlink to the choice table"
    Assert.AreEqual 1, dropStub.ReturnLinkCount, _
                     "Custom dropdowns should add a return link back to the variable label"
    Assert.AreEqual "test_var", dropStub.ValidationListName(dropStub.ValidationCallCount), _
                     "Custom dropdown validation should target the variable list"
    Assert.IsTrue dropStub.ValidationShowError(dropStub.ValidationCallCount), _
                  "Custom choice validation should surface errors for invalid selections"
    Assert.IsTrue design.ScopeCount(HListMainLab) >= 2, _
                  "Custom dropdown hyperlinks should receive the main label styling after Excel removes it"
    Assert.IsTrue design.ScopeCount(HListSublab) >= 2, _
                  "Custom dropdown hyperlinks should restore the sub label styling after Excel removes it"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsListAutoChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Const LIST_NAME As String = "auto_list"

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set context = BuildContext("text", controlType:="list_auto", controlDetails:=LIST_NAME, _
                               dropdownDefault:=dropStub)

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual LIST_NAME, dropStub.LastAddedListName, _
                     "List auto controls should register the dropdown using the control details name"
    Assert.AreEqual LIST_NAME, dropStub.ValidationListName(1), _
                     "List auto validation should reference the generated dropdown name"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsGeoChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim geoStub As LLGeoDropdownStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set geoStub = New LLGeoDropdownStub
    geoStub.SetLevel LevelAdmin1, Array("Region A", "Region B")

    Set context = BuildContext("text", controlType:="geo1", dropdownDefault:=dropStub)
    CurrentSpecsStub.SetGeo geoStub

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual 4, dropStub.ValidationCallCount, _
                     "Geo controls should wire four cascading validations"
    Assert.AreEqual layout.ValueCell.Address(False, False), dropStub.ValidationTargetAddress(1), _
                     "Admin 1 validation should target the primary value cell"
    Assert.AreEqual layout.ValueCell.Offset(0, 1).Address(False, False), dropStub.ValidationTargetAddress(2), _
                     "Admin 2 validation should target the next column"
    Assert.AreEqual "admin3", dropStub.ValidationListName(3), _
                     "Admin 3 validation should target the admin3 list"
    Assert.AreEqual "admin4", dropStub.ValidationListName(4), _
                     "Admin 4 validation should target the admin4 list"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAppliesFormulaControl()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim printSheet As Worksheet
    Dim design As LLFormatLogStub
    Dim dict As ILLdictionary
    Dim formulaData As FormulaDataStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear
    Set printSheet = TestHelpers.EnsureWorksheet(WRITER_PRINT_SHEET)
    printSheet.Cells.Clear

    Set design = New LLFormatLogStub
    Set dict = New DictionaryMinimalStub
    Set formulaData = New FormulaDataStub
    formulaData.AllowCharacter "+"

    Set context = BuildContext("decimal", controlType:="formula", controlDetails:="1+1", _
                               design:=design, dictionary:=dict, formulaData:=formulaData)
    CurrentLinelistStub.UsePrintedSheet context.ValueOf("sheet name"), printSheet

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3, printSheet
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "=1+1", layout.ValueCell.Formula, _
                     "Formula controls should translate the linelist expression into an Excel formula"
    Assert.IsTrue layout.ValueCell.Locked, _
                  "Calculated variables should lock the value cell"
    Assert.IsTrue layout.ValueCell.FormulaHidden, _
                  "Calculated variables should hide the source formula"
    Assert.IsTrue design.ScopeCount(HListCalculatedFormulaCell) >= 1, _
                  "Calculated variables should apply the calculated styling scope to the entry cell"
    Assert.IsTrue design.ScopeCount(HListCalculatedFormulaHeader) >= 2, _
                  "Calculated variables should format both linelist and printed headers with the calculated scope"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsFormattingCondition()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim templateRange As Range
    Dim condition As FormatCondition
    Dim flagRow As Long
    Dim expectedFormula As String

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("integer", formattingCondition:="flag_var")

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3

    flagRow = layout.ValueCell.Row + 5

    Set templateRange = sheet.Cells(2, layout.ValueCell.Column + 5)
    templateRange.Interior.Color = RGB(220, 130, 70)
    templateRange.Font.Color = RGB(30, 40, 90)
    templateRange.Font.Bold = True
    templateRange.Font.Italic = True

    sheet.Cells(flagRow, layout.ValueCell.Column).Value = 0

    CurrentVariablesStub.AddValue "flag_var", "sheet name", WRITER_SHEET
    CurrentVariablesStub.AddValue "flag_var", "column index", CStr(flagRow)
    CurrentVariablesStub.AddValue "flag_var", "variable name", "flag_var"
    CurrentVariablesStub.SetCellRange "test_var", "formatting values", templateRange

    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual 1, layout.ValueCell.FormatConditions.Count, _
                     "Formatting conditions should add a single conditional formatting rule to the value cell"

    Set condition = layout.ValueCell.FormatConditions(1)

    expectedFormula = "=('" & Replace(sheet.Name, "'", "''") & "'!" & _
                      sheet.Cells(flagRow, layout.ValueCell.Column).Address(RowAbsolute:=True, ColumnAbsolute:=True) & "=1)"

    Assert.AreEqual Replace(expectedFormula, " ", vbNullString), _
                     Replace(condition.Formula1, " ", vbNullString), _
                     "Conditional formatting must monitor the supplied formatting condition variable"
    Assert.AreEqual templateRange.Interior.Color, condition.Interior.Color, _
                     "Conditional formatting should adopt the interior colour stored in the formatting template"
    Assert.AreEqual templateRange.Font.Color, condition.Font.Color, _
                     "Conditional formatting should mirror the template font colour"
    Assert.IsTrue condition.Font.Bold, _
                  "Conditional formatting should mirror the template bold flag"
    Assert.IsTrue condition.Font.Italic, _
                  "Conditional formatting should mirror the template italic flag"
    Assert.IsTrue condition.StopIfTrue, _
                  "Formatting conditions should stop processing later rules when the flag evaluates to 1"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsUniqueCondition()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim condition As FormatCondition

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set context = BuildContext("text")
    CurrentVariablesStub.AddValue "test_var", "unique", "yes"

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual 1, layout.ValueCell.FormatConditions.Count, _
                     "Unique variables should receive a duplicate detection rule"

    Set condition = layout.ValueCell.FormatConditions(1)

    Assert.AreEqual xlDuplicate, condition.DupeUnique, _
                     "Unique formatting should mark duplicate values"
    Assert.AreEqual vbRed, condition.Interior.Color, _
                     "Duplicate highlighting should use a red fill to raise attention"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestHorizontalWriterAddsGeoWarning()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutHorizontal
    Dim hooks As LLVarWriterHooksHorizontal
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim geoStub As LLGeoDropdownStub
    Dim pcodeRow As Long
    Dim condition As FormatCondition
    Dim expectedFormula As String

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set geoStub = New LLGeoDropdownStub
    geoStub.SetLevel LevelAdmin1, Array("Zone 1")

    Set context = BuildContext("text", controlType:="geo1", dropdownDefault:=dropStub)
    CurrentSpecsStub.SetGeo geoStub

    Set layout = New LLVarLayoutHorizontal
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksHorizontal

    pcodeRow = layout.ValueCell.Row + 3
    sheet.Cells(pcodeRow, layout.ValueCell.Column).Value = vbNullString

    CurrentVariablesStub.AddValue "pcode_test_var", "sheet name", WRITER_SHEET
    CurrentVariablesStub.AddValue "pcode_test_var", "column index", CStr(pcodeRow)
    CurrentVariablesStub.AddValue "pcode_test_var", "variable name", "pcode_test_var"

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual 1, layout.ValueCell.FormatConditions.Count, _
                     "Geo controls should add a warning when the associated pcode is blank"

    Set condition = layout.ValueCell.FormatConditions(1)

    expectedFormula = "=AND('" & Replace(sheet.Name, "'", "''") & "'!" & _
                      sheet.Cells(pcodeRow, layout.ValueCell.Column).Address(RowAbsolute:=True, ColumnAbsolute:=True) & "=" & Chr$(34) & Chr$(34) & "," & _
                      "'" & Replace(sheet.Name, "'", "''") & "'!" & _
                      layout.ValueCell.Address(RowAbsolute:=True, ColumnAbsolute:=True) & "<>" & Chr$(34) & Chr$(34) & ")"

    Assert.AreEqual Replace(expectedFormula, " ", vbNullString), _
                     Replace(condition.Formula1, " ", vbNullString), _
                     "Geo warnings should highlight entries lacking the associated pcode"
    Assert.AreEqual RGB(242, 147, 12), condition.Interior.Color, _
                     "Geo warnings should use the legacy amber fill colour"
    Assert.IsTrue condition.StopIfTrue, _
                  "Geo warnings should short-circuit any later conditional formats"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterAppliesFormulaControl()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet
    Dim design As LLFormatLogStub
    Dim dict As ILLdictionary
    Dim formulaData As FormulaDataStub

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set design = New LLFormatLogStub
    Set dict = New DictionaryMinimalStub
    Set formulaData = New FormulaDataStub
    formulaData.AllowCharacter "+"

    Set context = BuildContext("decimal", controlType:="formula", controlDetails:="1+1", _
                               design:=design, dictionary:=dict, formulaData:=formulaData)
    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "=1+1", layout.ValueCell.Formula, _
                     "Formula controls should translate the linelist expression into an Excel formula"
    Assert.IsTrue layout.ValueCell.Locked, _
                  "Calculated variables should lock the value cell"
    Assert.AreEqual 1, design.ScopeCount(HListCalculatedFormulaCell), _
                     "Calculated variables should apply the calculated styling scope"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterAddsManualChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim categories As BetterArray

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set context = BuildContext("text", controlType:="choice_manual", dropdownDefault:=dropStub)
    Set categories = New BetterArray
    categories.LowerBound = 1
    categories.Push "Option A"
    categories.Push "Option B"
    CurrentSpecsStub.SetCategories "test_var", categories

    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual "test_var", dropStub.LastAddedListName, _
                     "Manual choices should register the dropdown by variable name"
    Assert.IsTrue dropStub.LastValidationShowError, _
                  "Manual lists should raise validation errors when entries are invalid"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterAddsCustomChoices()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet
    Dim dropStub As DropdownListsStub
    Dim categories As BetterArray
    Dim translation As LinelistSpecsTranslationStub
    Dim design As LLFormatLogStub
    Dim labelText As String

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set dropStub = New DropdownListsStub
    dropStub.Initialise sheet

    Set translation = New LinelistSpecsTranslationStub
    translation.Initialise
    translation.SetTranslation "MSG_CustomChoice", "Custom"

    Set design = New LLFormatLogStub

    Set context = BuildContext("text", controlType:="choice_custom", dropdownDefault:=dropStub, _
                               dropdownCustom:=dropStub, design:=design, translation:=translation)
    Set categories = New BetterArray
    categories.LowerBound = 1
    categories.Push "Custom Option"
    CurrentSpecsStub.SetCategories "test_var", categories

    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    labelText = CStr(layout.LabelCell.Value)

    Assert.IsTrue InStr(labelText, "(Custom") > 0, _
                  "Custom dropdowns should append the generated entry label to the sub label"
    Assert.AreEqual 1, layout.LabelCell.Hyperlinks.Count, _
                     "Custom dropdowns should create a hyperlink to the choice table"
    Assert.AreEqual 1, dropStub.ReturnLinkCount, _
                     "Custom dropdowns should add a return link back to the variable label"
    Assert.IsTrue dropStub.LastValidationShowError, _
                  "Custom dropdown validations should surface errors for invalid selections"
    Assert.AreEqual 2, design.ScopeCount(VListMainLab), _
                     "Custom dropdown hyperlinks should receive the main label styling after Excel removes it"
    Assert.AreEqual 2, design.ScopeCount(VListSublab), _
                     "Custom dropdown hyperlinks should restore the sub label styling after Excel removes it"
End Sub

'@TestMethod("LLVarWriter")
Private Sub TestVerticalWriterAddsFormattingCondition()
    Dim context As ILLVarContext
    Dim writer As LLVarWriterBase
    Dim layout As LLVarLayoutVertical
    Dim hooks As LLVarWriterHooksVertical
    Dim sheet As Worksheet
    Dim templateRange As Range
    Dim flagRow As Long
    Dim expectedFormula As String
    Dim condition As FormatCondition

    flagRow = 6

    Set sheet = TestHelpers.EnsureWorksheet(WRITER_SHEET)
    sheet.Cells.Clear

    Set templateRange = sheet.Cells(2, 10)
    templateRange.Interior.Color = RGB(220, 130, 70)
    templateRange.Font.Color = RGB(30, 40, 90)
    templateRange.Font.Bold = True
    templateRange.Font.Italic = True

    sheet.Cells(flagRow, 5).Value = 0

    Set context = BuildContext("integer", formattingCondition:="flag_var")
    CurrentVariablesStub.AddValue "flag_var", "sheet name", WRITER_SHEET
    CurrentVariablesStub.AddValue "flag_var", "column index", CStr(flagRow)
    CurrentVariablesStub.AddValue "flag_var", "variable name", "flag_var"
    CurrentVariablesStub.SetCellRange "test_var", "formatting values", templateRange

    Set layout = New LLVarLayoutVertical
    layout.Configure sheet, 3
    Set hooks = New LLVarWriterHooksVertical

    Set writer = New LLVarWriterBase
    writer.Initialise context, layout, hooks
    writer.WriteVariable

    Assert.AreEqual 1, layout.ValueCell.FormatConditions.Count, _
                     "Formatting conditions should add a single conditional formatting rule to the value cell"

    Set condition = layout.ValueCell.FormatConditions(1)

    expectedFormula = "=('"
    expectedFormula = expectedFormula & Replace(sheet.Name, "'", "''")
    expectedFormula = expectedFormula & "'!"
    expectedFormula = expectedFormula & sheet.Cells(flagRow, layout.ValueCell.Column).Address(RowAbsolute:=True, ColumnAbsolute:=True)
    expectedFormula = expectedFormula & "=1)"

    Assert.AreEqual Replace(expectedFormula, " ", vbNullString), _
                     Replace(condition.Formula1, " ", vbNullString), _
                     "Conditional formatting must monitor the supplied formatting condition variable"
    Assert.AreEqual templateRange.Interior.Color, condition.Interior.Color, _
                     "Conditional formatting should adopt the interior colour stored in the formatting template"
    Assert.AreEqual templateRange.Font.Color, condition.Font.Color, _
                     "Conditional formatting should mirror the template font colour"
    Assert.IsTrue condition.Font.Bold, _
                  "Conditional formatting should mirror the template bold flag"
    Assert.IsTrue condition.Font.Italic, _
                  "Conditional formatting should mirror the template italic flag"
    Assert.IsTrue condition.StopIfTrue, _
                  "Formatting conditions should stop processing later rules when the flag evaluates to 1"
End Sub
