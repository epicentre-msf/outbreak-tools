Attribute VB_Name = "TestLLVarWriterBase"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const WRITER_SHEET As String = "LLVarWriterSheet"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet WRITER_SHEET
    Set Assert = Nothing
End Sub

Private Function BuildContext(ByVal varType As String, _
                              Optional ByVal varFormat As String = vbNullString, _
                              Optional ByVal minValue As String = vbNullString, _
                              Optional ByVal maxValue As String = vbNullString) As ILLVarContext

    Dim linelistStub As LLVarContextLinelistStub
    Dim variablesStub As LLVarContextVariablesStub
    Dim specsStub As LLVarContextSpecsStub
    Dim context As ILLVarContext

    Set linelistStub = New LLVarContextLinelistStub
    linelistStub.UseWorkbook ThisWorkbook

    Set specsStub = New LLVarContextSpecsStub
    linelistStub.UseSpecs specsStub

    Set variablesStub = New LLVarContextVariablesStub
    variablesStub.AddValue "test_var", "sheet name", WRITER_SHEET
    variablesStub.AddValue "test_var", "column index", "3"
    variablesStub.AddValue "test_var", "main label", "Main"
    variablesStub.AddValue "test_var", "sub label", "Sub"
    variablesStub.AddValue "test_var", "variable name", "test_var"
    variablesStub.AddValue "test_var", "variable type", varType
    variablesStub.AddValue "test_var", "variable format", varFormat
    variablesStub.AddValue "test_var", "status", "active"
    variablesStub.AddValue "test_var", "note", "A helpful note"
    variablesStub.AddValue "test_var", "control", "text"
    variablesStub.AddValue "test_var", "min", minValue
    variablesStub.AddValue "test_var", "max", maxValue
    variablesStub.AddValue "test_var", "alert", "warning"
    variablesStub.AddValue "test_var", "message", "Range check"

    Set context = New LLVarContext
    context.Initialise linelistStub, "test_var", , variablesStub, specsStub

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

