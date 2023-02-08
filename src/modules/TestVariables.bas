Attribute VB_Name = "TestVariables"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private Dictionary As ILLdictionary
Private Variables As ILLVariables

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
    Set Variables = LLVariables.Create(Dictionary)
End Sub

'@TestMethod
Private Sub TestVariableValues()
    On Error GoTo VariableValuesFail
    
    Dim val As String
    
    val = Variables.Value(varName:="varb1", colName:="sheet type")
    Assert.IsTrue (val = "hlist2D"), "returned value of sheet type for variable varb1 is not correct. Expected hlist2D, returned : " & val
    
    val = Variables.Value(varName:="vara1", colName:="sheet type")
    Assert.IsTrue (val = "vlist1D"), "returned value of sheet type of variable vara1 is not correct Expected vlist1D, returned :" & val

    Assert.IsTrue Variables.Contains("varb1"), "varb1 exists as a variable, but it not found as one."
    Assert.IsFalse Variables.Contains("va"), "va does not exist as a variable, but it is found as one."
    Assert.IsFalse Variables.Contains(""), "empty characters are considered as present in variables"

    Exit Sub

VariableValuesFail:
    Assert.Fail "Variable values Failed: #" & Err.Number & " : " & Err.Description
End Sub

'@TestMethod
Private Sub TestIndex()
    Dim dataWksh As Worksheet
    Dim dict As ILLdictionary
    Dim vars As ILLVariables

    Set dataWksh = ThisWorkbook.Worksheets("Dictionary")
    Set dict = LLdictionary.Create(dataWksh, 1, 1)
    Dim sheetIndex As Long

    'TEst the column index
    dict.Prepare
    Set vars = LLVariables.Create(dict)
    sheetIndex = vars.Index("vara1")
    Assert.IsTrue (sheetIndex = 4), "Expected index: 4, Obtained index: " & sheetIndex
    sheetIndex = vars.Index("varb2")
    Assert.IsTrue (sheetIndex = 2), "Expected index: 2, Obtained index: " & sheetIndex
End Sub

