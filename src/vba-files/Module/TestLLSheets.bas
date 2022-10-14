Attribute VB_Name = "TestLLSheets"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private Dictionary As ILLdictionary
Private sheets As ILLSheets

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
    Set Dictionary = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Dim dataWksh As Worksheet
    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("TestDictionary")
    Set Dictionary = LLdictionary.Create(dataWksh, 1, 1)
    Dictionary.Prepare
    Set sheets = LLSheets.Create(Dictionary)
End Sub