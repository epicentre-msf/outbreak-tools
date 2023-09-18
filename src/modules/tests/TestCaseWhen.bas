Attribute VB_Name = "TestCaseWhen"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private casewhenObject As ICaseWhen
Private formula As String

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
    Set casewhenObject = Nothing
End Sub

'This method runs before every test in the module..

'@TestInitialize
Private Sub TestInitialize()
End Sub

Private Function Quoted(ByVal Text As String)
    Quoted = chr(34) & Text & chr(34)
End Function

'@TestMethod
Private Sub testcasewhen()
    Dim cat As BetterArray
    Dim dict As ILLdictionary
    Dim vars As ILLVariables

    On Error GoTo Fail
    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("TestDictionary"), 1, 1)
    Set vars = LLVariables.Create(dict)
    formula = vars.Value(varName:="vara4", colName:="control details")
    Set casewhenObject = CaseWhen.Create(formula)
    Assert.IsTrue (casewhenObject.Valid()), "Validity test failed to succeed"
    Assert.IsTrue (casewhenObject.parsedFormula = ThisWorkbook.Worksheets("TestValues").Range("case_when_value").Value), "Case when not parsed correctly"
    Set cat = casewhenObject.Categories()
    Assert.IsTrue (cat.Item(1) = "Choice is A" And cat.Item(2) = "Choice is B"), "Case when categories not correct"
    formula = "IF(CASE_WHEN(yes, true)"
    Set casewhenObject = CaseWhen.Create(formula)
    Assert.IsFalse casewhenObject.Valid(), "Validity test failed to fail"
    'parsed case when

    Exit Sub
Fail:
    Assert.Fail "Test Case when failed #" & Err.Number & " : " & Err.Description
End Sub
