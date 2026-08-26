Attribute VB_Name = "TestDiseaseComparison"
Attribute VB_Description = "Tests covering DiseaseComparison over a fixture pair of disease tables"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests covering DiseaseComparison over a fixture pair of disease tables")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const FIRST_SHEET_NAME As String = "DiseaseCompareFirst"
Private Const SECOND_SHEET_NAME As String = "DiseaseCompareSecond"
Private Const FIRST_TABLE_NAME As String = "T_CompareFirst"
Private Const SECOND_TABLE_NAME As String = "T_CompareSecond"

Private Assert As CustomTest
Private FirstSheet As Worksheet
Private SecondSheet As Worksheet
Private FirstTable As ListObject
Private SecondTable As ListObject

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDiseaseComparison"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    On Error Resume Next
        If Not Assert Is Nothing Then
            Assert.PrintResults TEST_OUTPUT_SHEET
        End If
        DeleteWorksheets FIRST_SHEET_NAME, SECOND_SHEET_NAME
    On Error GoTo 0

    RestoreApp
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    PrepareFirstTable
    PrepareSecondTable
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'The assertions of a test reach the results sheet only once flushed.
    Assert.Flush
    If Not FirstSheet Is Nothing Then ClearWorksheet FirstSheet
    If Not SecondSheet Is Nothing Then ClearWorksheet SecondSheet
    Set FirstTable = Nothing
    Set SecondTable = Nothing
End Sub

'@section Tests
'===============================================================================

'The fixture pair:
'  var_a  shared, same choices spelled in another order and case
'  var_b  only in the first disease
'  var_c  only in the second disease
'  var_d  shared, choices differ ("no" only in 1, "maybe" only in 2)
'  var_e  shared, Choice cell empty on both sides

'@TestMethod("DiseaseComparison")
Public Sub TestOnlyInFirstListsTheMissingVariables()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestOnlyInFirstListsTheMissingVariables"

    Dim comparison As DiseaseComparison
    Dim bundle As Checking
    Dim keys As BetterArray

    On Error GoTo Fail

    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable, "Cholera", "Measles")
    Set bundle = comparison.OnlyInFirst

    Assert.AreEqual "Variables only in Cholera", bundle.CheckingTitle, "The bundle names the first disease"
    Assert.AreEqual "Type--Where?--Details", bundle.CheckingSubTitle, "The bundle takes the report subtitle"
    Assert.AreEqual 1, bundle.Length, "One variable is only in the first disease"

    Set keys = bundle.ListOfKeys
    Assert.AreEqual "only-in-1-1", CStr(keys.Item(keys.LowerBound)), "The entry key carries the section prefix"
    Assert.IsTrue InStr(1, bundle.ValueOf("only-in-1-1"), "var_b (symptoms)", vbTextCompare) > 0, _
                  "The entry names the variable and its section"
    Assert.IsTrue InStr(1, bundle.ValueOf("only-in-1-1"), "Label: LabelB", vbTextCompare) > 0, _
                  "The entry carries the label"
    Assert.IsTrue InStr(1, bundle.ValueOf("only-in-1-1"), "Missing from Measles", vbTextCompare) > 0, _
                  "The entry names the disease the variable is missing from"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestOnlyInFirstListsTheMissingVariables", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestOnlyInSecondListsTheMirror()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestOnlyInSecondListsTheMirror"

    Dim comparison As DiseaseComparison
    Dim bundle As Checking

    On Error GoTo Fail

    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable, "Cholera", "Measles")
    Set bundle = comparison.OnlyInSecond

    Assert.AreEqual "Variables only in Measles", bundle.CheckingTitle, "The bundle names the second disease"
    Assert.AreEqual 1, bundle.Length, "One variable is only in the second disease"
    Assert.IsTrue bundle.KeyExists("only-in-2-1"), "The entry key carries the section prefix"
    Assert.IsTrue InStr(1, bundle.ValueOf("only-in-2-1"), "var_c (history)", vbTextCompare) > 0, _
                  "The entry names the variable and its section"
    Assert.IsTrue InStr(1, bundle.ValueOf("only-in-2-1"), "Missing from Cholera", vbTextCompare) > 0, _
                  "The entry names the first disease as the one missing it"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestOnlyInSecondListsTheMirror", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestDifferingChoicesListsEachMissingValue()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestDifferingChoicesListsEachMissingValue"

    Dim comparison As DiseaseComparison
    Dim bundle As Checking

    On Error GoTo Fail

    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable, "Cholera", "Measles")
    Set bundle = comparison.DifferingChoices

    Assert.AreEqual "Shared variables whose choices differ", bundle.CheckingTitle, "The bundle carries its title"
    Assert.AreEqual 3, bundle.Length, "One entry names the variable, one per value on each side"

    Assert.IsTrue bundle.KeyExists("differing-1"), "The first entry names the variable"
    Assert.IsTrue InStr(1, bundle.ValueOf("differing-1"), "var_d (outcome)", vbTextCompare) > 0, _
                  "The variable entry names the variable and its section"

    Assert.IsTrue bundle.KeyExists("differing-1-first-1"), "One value is in the first disease only"
    Assert.IsTrue InStr(1, bundle.ValueOf("differing-1-first-1"), """no""", vbTextCompare) > 0, _
                  "The value only in the first disease is named"
    Assert.IsTrue InStr(1, bundle.ValueOf("differing-1-first-1"), "missing from Measles", vbTextCompare) > 0, _
                  "The value entry says which disease lacks it"

    Assert.IsTrue bundle.KeyExists("differing-1-second-1"), "One value is in the second disease only"
    Assert.IsTrue InStr(1, bundle.ValueOf("differing-1-second-1"), """maybe""", vbTextCompare) > 0, _
                  "The value only in the second disease is named"

    Assert.AreEqual "Warning", TypeWord(bundle.ValueOf("differing-1", checkingType)), _
                    "A choice-set difference is a warning"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestDifferingChoicesListsEachMissingValue", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestEquivalentVariablesIgnoreOrderCaseAndEmptyChoice()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestEquivalentVariablesIgnoreOrderCaseAndEmptyChoice"

    Dim comparison As DiseaseComparison
    Dim bundle As Checking

    On Error GoTo Fail

    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable, "Cholera", "Measles")
    Set bundle = comparison.EquivalentVariables

    Assert.AreEqual "Equivalent variables", bundle.CheckingTitle, "The bundle carries its title"
    Assert.AreEqual 2, bundle.Length, "var_a and var_e are equivalent"
    Assert.IsTrue InStr(1, bundle.ValueOf("equivalent-1"), "var_a (demographics)", vbTextCompare) > 0, _
                  "Choices in another order and case still match"
    Assert.IsTrue InStr(1, bundle.ValueOf("equivalent-2"), "var_e (notes)", vbTextCompare) > 0, _
                  "Two variables with an empty Choice cell are equivalent"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestEquivalentVariablesIgnoreOrderCaseAndEmptyChoice", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestStatisticsAnswerTheCounts()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestStatisticsAnswerTheCounts"

    Dim comparison As DiseaseComparison
    Dim bundle As Checking

    On Error GoTo Fail

    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable, "Cholera", "Measles")

    Assert.AreEqual 4, comparison.FirstCount, "The first disease carries four variables"
    Assert.AreEqual 4, comparison.SecondCount, "The second disease carries four variables"
    Assert.AreEqual 3, comparison.SharedCount, "Three names are shared"
    Assert.AreEqual 1, comparison.OnlyInFirstCount, "One variable is only in the first disease"
    Assert.AreEqual 1, comparison.OnlyInSecondCount, "One variable is only in the second disease"
    Assert.AreEqual 2, comparison.MatchingCount, "Two shared variables have matching choices"
    Assert.AreEqual 1, comparison.DifferingCount, "One shared variable has differing choices"
    Assert.IsTrue Abs(comparison.EquivalentShare - (2 / 3)) < 0.0001, "Two thirds of the shared variables are equivalent"

    Set bundle = comparison.Statistics
    Assert.AreEqual "Comparison statistics", bundle.CheckingTitle, "The statistics bundle carries its title"
    Assert.AreEqual 8, bundle.Length, "The statistics bundle carries eight lines"
    Assert.IsTrue InStr(1, bundle.ValueOf("stat-shared"), "3 shared", vbTextCompare) > 0, "The shared count is written"
    Assert.IsTrue InStr(1, bundle.ValueOf("stat-share"), "66.7%", vbTextCompare) > 0, "The share is written as a percentage"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestStatisticsAnswerTheCounts", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestReportsAnswerTheFiveBundlesInOrder()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestReportsAnswerTheFiveBundlesInOrder"

    Dim comparison As DiseaseComparison
    Dim bundles As BetterArray
    Dim bundle As Checking

    On Error GoTo Fail

    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable)
    Set bundles = comparison.Reports

    Assert.AreEqual 1, bundles.LowerBound, "The bundle list is based at 1"
    Assert.AreEqual 5, bundles.Length, "Five bundles are answered"

    Set bundle = bundles.Item(1)
    Assert.AreEqual "Variables only in Disease 1", bundle.CheckingTitle, "The default name of the first disease is used"
    Set bundle = bundles.Item(2)
    Assert.AreEqual "Variables only in Disease 2", bundle.CheckingTitle, "The default name of the second disease is used"
    Set bundle = bundles.Item(3)
    Assert.AreEqual "Shared variables whose choices differ", bundle.CheckingTitle, "The third bundle is the differing one"
    Set bundle = bundles.Item(4)
    Assert.AreEqual "Equivalent variables", bundle.CheckingTitle, "The fourth bundle is the equivalent one"
    Set bundle = bundles.Item(5)
    Assert.AreEqual "Comparison statistics", bundle.CheckingTitle, "The fifth bundle is the statistics"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestReportsAnswerTheFiveBundlesInOrder", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestCreateRefusesAMissingTable()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestCreateRefusesAMissingTable"

    Dim comparison As DiseaseComparison
    Dim raisedNumber As Long

    On Error GoTo Fail

    On Error Resume Next
        Set comparison = DiseaseComparison.Create(FirstTable, Nothing)
        raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), raisedNumber, "A Nothing table raises ObjectNotInitialized"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateRefusesAMissingTable", Err.Number, Err.Description
End Sub

'@TestMethod("DiseaseComparison")
Public Sub TestCompareRefusesATableWithoutTheNameColumn()
    CustomTestSetTitles Assert, "DiseaseComparison", "TestCompareRefusesATableWithoutTheNameColumn"

    Dim comparison As DiseaseComparison
    Dim raisedNumber As Long

    On Error GoTo Fail

    SecondTable.ListColumns("Variable Name").Delete
    Set comparison = DiseaseComparison.Create(FirstTable, SecondTable)

    On Error Resume Next
        comparison.Compare
        raisedNumber = Err.Number
    On Error GoTo Fail

    Assert.AreEqual CLng(ProjectError.ElementNotFound), raisedNumber, "A table without the name column raises ElementNotFound"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCompareRefusesATableWithoutTheNameColumn", Err.Number, Err.Description
End Sub

'@section Fixtures
'===============================================================================

Private Sub PrepareFirstTable()
    Dim headerMatrix As Variant
    Dim bodyMatrix As Variant
    Dim tableRange As Range

    Set FirstSheet = EnsureWorksheet(FIRST_SHEET_NAME)
    ClearWorksheet FirstSheet

    headerMatrix = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    bodyMatrix = RowsToMatrix(Array( _
        Array(1, "var_a", "demographics", "LabelA", "choiceA", "0-4 | 5-14 | 15+", "core"), _
        Array(2, "var_b", "symptoms", "LabelB", "choiceB", "yes | no", "core"), _
        Array(3, "var_d", "outcome", "LabelD", "choiceD", "yes | no", "core"), _
        Array(4, "var_e", "notes", "LabelE", "", "", "optional") _
    ))

    WriteMatrix FirstSheet.Range("A1"), headerMatrix
    WriteMatrix FirstSheet.Range("A2"), bodyMatrix

    Set tableRange = FirstSheet.Range("A1").Resize(UBound(bodyMatrix, 1) + 1, 7)
    Set FirstTable = FirstSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                XlListObjectHasHeaders:=xlYes)
    FirstTable.Name = FIRST_TABLE_NAME
End Sub

Private Sub PrepareSecondTable()
    Dim headerMatrix As Variant
    Dim bodyMatrix As Variant
    Dim tableRange As Range

    Set SecondSheet = EnsureWorksheet(SECOND_SHEET_NAME)
    ClearWorksheet SecondSheet

    headerMatrix = RowsToMatrix(Array(Array("Variable Order", "Variable Name", "Variable Section", "Main Label", "Choice", "Choice Values", "Status")))
    bodyMatrix = RowsToMatrix(Array( _
        Array(1, "VAR_A", "demographics", "LabelA", "choiceA", "15+ | 5-14 | 0-4", "core"), _
        Array(2, "var_c", "history", "LabelC", "choiceC", "low | high", "optional"), _
        Array(3, "var_d", "outcome", "LabelD", "choiceD", "yes | maybe", "core"), _
        Array(4, "var_e", "notes", "LabelE", "", "stale | values", "optional") _
    ))

    WriteMatrix SecondSheet.Range("A1"), headerMatrix
    WriteMatrix SecondSheet.Range("A2"), bodyMatrix

    Set tableRange = SecondSheet.Range("A1").Resize(UBound(bodyMatrix, 1) + 1, 7)
    Set SecondTable = SecondSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, _
                                                  XlListObjectHasHeaders:=xlYes)
    SecondTable.Name = SECOND_TABLE_NAME
End Sub

'The type descriptor of Checking carries an icon before the word; the word
'is what a test compares.
Private Function TypeWord(ByVal descriptor As String) As String
    Dim pieces() As String
    pieces = Split(Trim$(descriptor), " ")
    TypeWord = pieces(UBound(pieces))
End Function
