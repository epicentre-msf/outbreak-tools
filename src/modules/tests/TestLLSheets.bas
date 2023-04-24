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
    Set sheets = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Dim dataWksh As Worksheet
    'This method runs before every test in the module..
    Set dataWksh = ThisWorkbook.Worksheets("Dictionary")
    Set Dictionary = LLdictionary.Create(dataWksh, 1, 1)
    Dictionary.Prepare
    Set sheets = LLSheets.Create(Dictionary)
End Sub

'Testing the initialisation
'@TestMethod
Private Sub TestSheetInit()

    Assert.IsTrue (sheets.Dictionary.Data.Wksh.Name = "Dictionary"), "Bad dictionary associated with the worksheet"
    Assert.IsTrue (sheets.Dictionary.Data.StartRow = 1), "Start row of the dictionary associated with the sheets object is not correct"
    Assert.IsTrue (sheets.Dictionary.Data.StartColumn = 1), "End row of the dictionary associated with the sheets object is not correct"

End Sub

'Testing the sheet info
'@TestMethod
Private Sub TestSheetInfo()
    On Error GoTo Fail
    Assert.IsTrue (sheets.sheetInfo("A, B, C") = "vlist1D"), "A vlist1D worksheet is not detected correctly in sheet Info"
    Assert.IsTrue (sheets.sheetInfo("A, B, C", 2) = "tab1"), "Unable to get the table name of a worksheet in sheet Info"

    Exit Sub
Fail:
    Assert.Fail "Sheet Info Failed: #" & Err.Number & " : " & Err.Description

End Sub

'Testing the sheet info
'@TestMethod
Private Sub TestSheetDataBounds()
    On Error GoTo Fail

    'Start
    Assert.IsTrue (sheets.DataBounds("A, B, C", 1) = 4), "Bad row start for sheet A, B, C"
    Assert.IsTrue (sheets.DataBounds("A, B, C", 3) = 5), "Bad col start for sheet A, B, C"
    Assert.IsTrue (sheets.DataBounds("B-H2D", 1) = 9), "Bad row start for sheet B-H2D"
    Assert.IsTrue (sheets.DataBounds("B-H2D", 3) = 1), "Bad col start for sheet B-H2D"

    'End
    Assert.IsTrue (sheets.DataBounds("A, B, C", 2) = 19), "Bad row end for sheet A, B, C"
    Assert.IsTrue (sheets.DataBounds("A, B, C", 4) = 5), "Bad col end for sheet A, B, C"
    Assert.IsTrue (sheets.DataBounds("B-H2D", 2) = 209), "Bad row end for sheet B-H2D"
    Assert.IsTrue (sheets.DataBounds("B-H2D", 4) = 29), "Bad col end for sheet B-H2D. " & "Expected 29, obtained: " & sheets.DataBounds("B-H2D", 4)

    Exit Sub
Fail:
    Assert.Fail "Sheet Info Failed: #" & Err.Number & " : " & Err.Description

End Sub

'Testing the sheet info
'@TestMethod
Private Sub TestSheetContains()
    On Error GoTo Fail

    'found sheet name
    Assert.IsTrue sheets.Contains("A, B, C"), "sheet A, B, C is available, but seen as not available"

    'unfound sheet name
    Assert.IsFalse sheets.Contains(""), "Empty sheet name exists"
    Assert.IsFalse sheets.Contains("mjkmdjlqsjfs"), "sheet does not exist, but found as present"

    Exit Sub
Fail:
    Assert.Fail "Sheet Contains failed: #" & Err.Number & " : " & Err.Description
End Sub

'Testing the sheet info
'@TestMethod
Private Sub TestSheetContainsVarsOf()
    On Error GoTo Fail

    'Contains list auto
    Assert.IsTrue sheets.ContainsVarsOf("B-H2D"), "List auto not found in sheet"
    Assert.IsFalse sheets.ContainsVarsOf("A, B, C"), "List auto does not exists, but Test found one"

    Exit Sub
Fail:
    Assert.Fail "Sheet list auto failed: #" & Err.Number & " : " & Err.Description
End Sub

'Testing the sheet info
'@TestMethod
Private Sub TestSheetVariableAddress()
    On Error GoTo Fail

    'Address on a H2D worksheet
    Assert.IsTrue sheets.VariableAddress("vara1") = "'A, B, C'!$E$4", "Vlist1D variable address is not correct"
    Assert.IsTrue sheets.VariableAddress("varb2") = "'B-H2D'!$B10", "hlist 2D variable address is not correct"

    'Address on a V1D worksheet
    Exit Sub
Fail:
    Assert.Fail "Variable Address failed #" & Err.Number & " : " & Err.Description
End Sub

