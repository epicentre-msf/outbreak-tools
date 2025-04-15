Attribute VB_Name = "TestLinelistSpecs"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private lData As ILinelistSpecs
Private dict As ILLdictionary
Private choi As ILLChoices

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
    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("TestDictionary"), 1, 1)
    Set choi = LLChoices.Create(ThisWorkbook.Worksheets("TestChoices"), 1, 1)
 
    Set lData = LinelistSpecs.Create(dict, choi)
End Sub

'@TestMethod
Private Sub TestPrepare()
    Dim cat As BetterArray
 
    lData.Prepare
    Set cat = choi.Categories("__case_when_vara4")
    Assert.IsTrue (cat.Length > 0), "case when categories not defined on vara4"
 
End Sub
