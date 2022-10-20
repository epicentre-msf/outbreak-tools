Attribute VB_Name = "TestFormulas"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private formData As IFormulaData
Private dict As ILLdictionary
Private llform As Formulas
Private parsedFormula As String
Private formCond As IFormulaCondition

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'This method runs before every test in the module..

'@TestInitialize
Private Sub TestInitialize()
    Dim Wksh As Worksheet
    Set Wksh = ThisWorkbook.Worksheets("ControleFormule")
    Set formData = FormulaData.Create(Wksh, "T_XlsFonctions", "T_ascii")
End Sub

'@TestMethod
Private Sub TestIncludes()
    On Error GoTo Fail

    Assert.IsTrue formData.SpecialCharacterIncludes("("), "Special character not found"
    Assert.IsTrue formData.ExcelFormulasIncludes("AVERAGE"), "Existing formula not found"
    Assert.IsFalse formData.ExcelFormulasIncludes("COMPLEXES"), "Non Existing formula found"

Exit Sub
Fail:
    Assert.Fail "Test Special characters failed #" & Err.Number & " : " & Err.Description

End Sub

