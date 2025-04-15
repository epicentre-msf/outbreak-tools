Attribute VB_Name = "TestTablesSpecs"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private lData As ILinelistSpecs
Private specs As ITablesSpecs

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
    Set specs = Nothing
    Set lData = Nothing
End Sub

'This method runs before every test in the module..
'@TestInitialize
Private Sub TestInitialize()
    Dim dict As ILLdictionary
    Dim choi As ILLChoices
    Dim headRng As Range
    Dim rowRng As Range
    Dim Lo As ListObject

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets("TestDictionary"), 1, 1)
    Set choi = LLChoices.Create(ThisWorkbook.Worksheets("TestChoices"), 1, 1)
    Set lData = LinelistSpecs.Create(dict, choi)
    Set Lo = ThisWorkbook.Worksheets("Analysis").ListObjects(2)
    Set headRng = Lo.HeaderRowRange
    Set rowRng = Lo.ListRows(1).Range
    Set specs = TablesSpecs.Create(headRng, rowRng, lData, TypeUnivariate)
End Sub

'Test row categories
'@TestMethod
Private Sub TestCategories()

    On Error GoTo Fail

    Dim cat As BetterArray
    Set cat = specs.RowCategories()

    Assert.IsTrue (cat.Item(1) = "A"), "Unable to find row categories in table specifications"

    Set cat = specs.ColumnCategories()
    Assert.IsTrue (cat.Length = 0), "Found unexisting column categories in table specifications"

    Exit Sub
Fail:
    Assert.Fail "Row Categories Failed: #" & Err.Number & " : " & Err.Description
End Sub

'Test value
'@TestMethod
Private Sub TestValueSection()
    On Error GoTo Fail

    Assert.IsTrue (specs.Value("section") = "Tables in section 1"), "Section value not found"
    Assert.IsTrue (specs.Value("graph") = "yes"), "Graph not found"
    Assert.IsTrue (specs.isNewSection), "New section not detected"
    Assert.IsTrue (specs.TableType = TypeUnivariate), "Bad table type for specs"

    Exit Sub
Fail:
    Assert.Fail "Row Categories Failed: #" & Err.Number & " : " & Err.Description
End Sub
