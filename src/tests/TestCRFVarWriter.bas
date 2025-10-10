Attribute VB_Name = "TestCRFVarWriter"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object

Private Const CRF_SHEET As String = "CRFWriterSheet"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet CRF_SHEET
    Set Assert = Nothing
End Sub

Private Function BuildCRFContext(ByVal controlType As String, _
                                 ByVal statusValue As String, _
                                 Optional ByVal crfChoices As String = "no", _
                                 Optional ByVal categories As Variant, _
                                 Optional ByVal shortCategories As Variant) As ILLVarContext

    Dim linelistStub As LLVarContextLinelistStub
    Dim variablesStub As LLVarContextVariablesStub
    Dim specsStub As LLVarContextSpecsStub
    Dim context As ILLVarContext

    TestHelpers.DeleteWorksheet CRF_SHEET

    Set linelistStub = New LLVarContextLinelistStub
    linelistStub.UseWorkbook ThisWorkbook

    Set specsStub = New LLVarContextSpecsStub
    If Not IsMissing(categories) Then
        specsStub.SetCategories "crf_var", categories
    End If
    If Not IsMissing(shortCategories) Then
        specsStub.SetCategories "crf_var", shortCategories, True
    End If
    linelistStub.UseSpecs specsStub

    Set variablesStub = New LLVarContextVariablesStub
    variablesStub.AddValue "crf_var", "sheet name", CRF_SHEET
    variablesStub.AddValue "crf_var", "table name", "crf_table"
    variablesStub.AddValue "crf_var", "crf index", "6"
    variablesStub.AddValue "crf_var", "main label", "CRF Main"
    variablesStub.AddValue "crf_var", "sub label", "CRF Sub"
    variablesStub.AddValue "crf_var", "variable name", "crf_var"
    variablesStub.AddValue "crf_var", "variable type", "text"
    variablesStub.AddValue "crf_var", "variable format", vbNullString
    variablesStub.AddValue "crf_var", "status", statusValue
    variablesStub.AddValue "crf_var", "note", "Important note"
    variablesStub.AddValue "crf_var", "control", controlType
    variablesStub.AddValue "crf_var", "crf choices", crfChoices

    Set context = New LLVarContext
    context.Initialise linelistStub, "crf_var", , variablesStub, specsStub

    Set BuildCRFContext = context
End Function

Private Sub WriteCRF(ByVal context As ILLVarContext)
    Dim writerFactory As New CRFVarWriter
    Dim writer As ILLVarWriter

    Set writer = writerFactory.Create(context)
    writer.WriteVariable
End Sub

'@TestMethod("CRFVarWriter")
Private Sub TestCRFWriterCreatesWorksheet()
    Dim context As ILLVarContext
    Dim sheet As Worksheet

    Set context = BuildCRFContext("text", "active")

    WriteCRF context

    Set sheet = ThisWorkbook.Worksheets(CRF_SHEET)
    Assert.AreEqual "CRF Main" & vbLf & "CRF Sub", sheet.Cells(6, 1).Value, _
                     "CRF writer should populate the label cell"
    Assert.AreEqual vbNullString, sheet.Cells(6, 2).Value, _
                     "CRF writer should leave adjacent cells empty unless required"
End Sub

'@TestMethod("CRFVarWriter")
Private Sub TestCRFWriterHidesRowsForHiddenStatus()
    Dim context As ILLVarContext
    Dim sheet As Worksheet

    Set context = BuildCRFContext("text", "hidden")
    WriteCRF context

    Set sheet = ThisWorkbook.Worksheets(CRF_SHEET)
    Assert.IsTrue sheet.Rows(6).Hidden, _
                  "Rows flagged as hidden in the dictionary should be hidden on the CRF sheet"
End Sub

'@TestMethod("CRFVarWriter")
Private Sub TestCRFWriterAddsChoiceColumns()
    Dim context As ILLVarContext
    Dim sheet As Worksheet

    Dim longCategories As Variant
    Dim shortCategories As Variant

    longCategories = Array("Long Yes", "Long No")
    shortCategories = Array("Y", "N")

    Set context = BuildCRFContext("choice_manual", "active", "yes", longCategories, shortCategories)

    WriteCRF context

    Set sheet = ThisWorkbook.Worksheets(CRF_SHEET)
    Assert.AreEqual "Y", sheet.Cells(4, 2).Value, _
                     "Choice headers should use the short label variant two rows above the CRF line"
    Assert.AreEqual "N", sheet.Cells(4, 4).Value, _
                     "Choice headers should advance two columns per entry when crf choices = yes"
    Assert.AreEqual "", sheet.Cells(6, 2).Value, _
                     "Choice value column should start empty"
    Assert.AreEqual "", sheet.Cells(6, 4).Value, _
                     "Second choice column should be created"
End Sub
