Attribute VB_Name = "TestLLVarContext"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private LinelistStub As ILinelist

Private Const CONTEXT_SHEET As String = "LLVarContextSheet"

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set LinelistStub = New LLVarContextLinelistStub
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    TestHelpers.DeleteWorksheet CONTEXT_SHEET
    Set LinelistStub = Nothing
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    TestHelpers.EnsureWorksheet CONTEXT_SHEET
End Sub

'@TestCleanup
Private Sub TestCleanup()
    TestHelpers.DeleteWorksheet CONTEXT_SHEET
End Sub

'@TestMethod("LLVarContext")
Private Sub TestContextInitialiseRequiresVariableName()
    Dim context As ILLVarContext
    Set context = New LLVarContext

    On Error Resume Next
        context.Initialise LinelistStub, vbNullString
        Dim errNumber As Long
        errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual ProjectError.InvalidArgument, errNumber, _
                     "Initialise should reject missing variable names"
End Sub

'@TestMethod("LLVarContext")
Private Sub TestContextValueOfReturnsDictionaryValues()
    Dim context As ILLVarContext
    Dim variablesStub As LLVarContextVariablesStub

    Set context = New LLVarContext
    Set variablesStub = New LLVarContextVariablesStub

    variablesStub.AddValue "var_one", "sheet name", CONTEXT_SHEET
    variablesStub.AddValue "var_one", "control", "text"

    context.Initialise LinelistStub, "var_one", , variablesStub

    Assert.AreEqual CONTEXT_SHEET, context.ValueOf("sheet name"), _
                     "Context should forward metadata lookups to the variables helper"
    Assert.AreEqual "text", context.ValueOf("control"), _
                     "Context should surface additional metadata columns"
End Sub

'@TestMethod("LLVarContext")
Private Sub TestContextWorksheetResolvesUsingMetadata()
    Dim context As ILLVarContext
    Dim variablesStub As LLVarContextVariablesStub
    Dim targetSheet As Worksheet

    Set context = New LLVarContext
    Set variablesStub = New LLVarContextVariablesStub

    variablesStub.AddValue "var_sheet", "sheet name", CONTEXT_SHEET

    context.Initialise LinelistStub, "var_sheet", , variablesStub

    Set targetSheet = context.Worksheet

    Assert.IsFalse targetSheet Is Nothing, _
                   "Worksheet lookup should return the target sheet"
    Assert.AreEqual CONTEXT_SHEET, targetSheet.Name, _
                     "Worksheet lookup should honour the dictionary metadata"

    'Request the worksheet again to ensure cached value is reused without error.
    Set targetSheet = context.Worksheet
    Assert.AreEqual CONTEXT_SHEET, targetSheet.Name, _
                     "Repeated worksheet lookups should return the cached reference"
End Sub

