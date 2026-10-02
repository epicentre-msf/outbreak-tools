Attribute VB_Name = "TestFormulaBuilder"
Attribute VB_Description = "Tests for FormulaBuilder class"
Option Explicit

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for FormulaBuilder class")

'@description
'Drives FormulaBuilder, the class CrossTableFormula and SpatialTables both build
'their formulas through. The tests cover the criteria set and the rule that a
'build closes it, the four text helpers, the write into a cell with Excel as the
'judge, the long array formula, and the shape of the keys the entries are filed
'under.
'@depends FormulaBuilder, Formulas, FormulaData, LLdictionary, Checking, CustomTest

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SCRATCH_SHEET As String = "FBuilderScratch"
Private Const DICT_SHEET As String = "FBuilderDict"
Private Const TOKENS_SHEET As String = "FBuilderTokens"
Private Const SCRATCH_BLOCK As String = "A1:H20"

Private Const OWNER_NAME As String = "TestOwner"
Private Const CONTEXT_ID As String = "T1"

Private Const ROW_CHOICE_VARIABLE As String = "choi_v1"
Private Const COUNT_CALL_FUNCTION As String = "N()"

' Range.FormulaArray refuses a formula longer than this.
Private Const FORMULA_ARRAY_LIMIT As Long = 255

Private Assert As CustomTest
Private dict As LLdictionary
Private fData As FormulaData

'@section Fixture helpers
'===============================================================================

'@sub-title Free a ListObject name wherever it is taken in the workbook.
'@param tableName String. The ListObject name to free.
Private Sub ReleaseTableName(ByVal tableName As String)
    Dim sh As Worksheet
    Dim idx As Long

    For Each sh In ThisWorkbook.Worksheets
        For idx = sh.ListObjects.Count To 1 Step -1
            If StrComp(sh.ListObjects(idx).Name, tableName, vbTextCompare) = 0 Then
                sh.ListObjects(idx).Unlist
            End If
        Next idx
    Next sh
End Sub

'@sub-title Return the scratch worksheet with nothing on it.
Private Function ScratchSheet() As Worksheet
    Dim sh As Worksheet

    Set sh = EnsureWorksheet(SCRATCH_SHEET, clearSheet:=False, visibility:=xlSheetHidden)
    sh.Range(SCRATCH_BLOCK).Clear
    Set ScratchSheet = sh
End Function

'@sub-title A builder filing into a report the test can read.
'@param checks Checking. ByRef. Filled with the report the builder writes into.
Private Function NewBuilder(ByRef checks As Checking) As FormulaBuilder
    Set checks = Checking.Create(OWNER_NAME)
    Set NewBuilder = FormulaBuilder.Create(OWNER_NAME, CONTEXT_ID, checks)
End Function

'@sub-title The messages a report holds, joined for a failure message.
'@param checks Checking. The report to read.
Private Function Messages(ByVal checks As Checking) As String
    Dim keyList As BetterArray
    Dim idx As Long
    Dim joined As String

    If checks.Length() <= 0 Then Exit Function

    Set keyList = checks.ListOfKeys

    For idx = keyList.LowerBound To keyList.UpperBound
        joined = joined & "[" & CStr(keyList.Item(idx)) & ": " & _
                 checks.ValueOf(CStr(keyList.Item(idx)), checkingLabel) & "]"
    Next idx

    Messages = joined
End Function

'@sub-title A formula text longer than the array formula limit.
'@details
'A sum of ones. Any Excel accepts it and computes it, so the answer in the cell
'is what proves the write went in.
Private Function LongFormulaText() As String
    Dim parts As String

    Do While Len(parts) <= FORMULA_ARRAY_LIMIT + 20
        If Len(parts) > 0 Then parts = parts & "+"
        parts = parts & "1"
    Loop

    LongFormulaText = parts
End Function

'@section Module lifecycle
'===============================================================================

'@sub-title Build the dictionary and the token tables.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestFormulaBuilder"

    PrepareDictionaryFixture DICT_SHEET
    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets(DICT_SHEET), 1, 1)
    dict.Prepare

    ' FormulaData resolves its two lookup tables by fixed names, so another suite
    ' fixture sheet holding them blocks this one from taking them.
    ReleaseTableName "T_XlsFonctions"
    ReleaseTableName "T_ascii"
    Set fData = FormulaData.Create(PrepareFormulaFixtureSheet(TOKENS_SHEET))
End Sub

'@sub-title Print results and tear down the fixtures.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    DeleteWorksheet SCRATCH_SHEET
    DeleteWorksheet DICT_SHEET
    DeleteWorksheet TOKENS_SHEET
    RestoreApp

    Set dict = Nothing
    Set fData = Nothing
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updating before each test.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush assert state after each test.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub

'@section Factory
'===============================================================================

'@sub-title Verify Create rejects a report that was not given.
'@TestMethod("FormulaBuilder")
Public Sub TestCreateRejectsNothingCheckings()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestCreateRejectsNothingCheckings"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim errNumber As Long

    On Error Resume Next
    Set builder = FormulaBuilder.Create(OWNER_NAME, CONTEXT_ID, Nothing)
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), errNumber, _
                    "A builder with nowhere to file its entries should be refused"
    Assert.IsTrue (builder Is Nothing), "Nothing should come back from a rejected Create"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestCreateRejectsNothingCheckings", Err.Number, Err.Description
End Sub

'@sub-title Verify the owner name and the table cannot be swapped after creation.
'@TestMethod("FormulaBuilder")
Public Sub TestTheOwnerIsSetAtCreationOnly()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestTheOwnerIsSetAtCreationOnly"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim errNumber As Long

    Set builder = NewBuilder(checks)

    On Error Resume Next
    builder.ContextId = "T2"
    errNumber = Err.Number
    On Error GoTo 0

    Assert.AreEqual CLng(ProjectError.SomethingWentWrong), errNumber, _
                    "Assigning the table after creation should raise"
    Assert.AreEqual CONTEXT_ID, builder.ContextId, "The table it was created for stands"
    Assert.AreEqual OWNER_NAME, builder.OwnerName, "The owner it was created for stands"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheOwnerIsSetAtCreationOnly", Err.Number, Err.Description
End Sub

'@section The criteria set
'===============================================================================

'@sub-title Verify each criterion added is counted.
'@TestMethod("FormulaBuilder")
Public Sub TestEachCriterionIsCounted()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestEachCriterionIsCounted"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking

    Set builder = NewBuilder(checks)

    Assert.AreEqual CLng(0), builder.ConditionCount, "A new builder holds no criteria"

    builder.AddCondition ROW_CHOICE_VARIABLE, "= $A$3"
    builder.AddCondition "int_v1", "<>" & builder.EmptyText

    Assert.AreEqual CLng(2), builder.ConditionCount, "Two criteria give a count of two"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEachCriterionIsCounted", Err.Number, Err.Description
End Sub

'@sub-title Verify a criterion with no variable name is ignored.
'@details
'That is how a caller passes a column the table does not group by.
'@TestMethod("FormulaBuilder")
Public Sub TestACriterionWithNoVariableIsIgnored()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestACriterionWithNoVariableIsIgnored"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking

    Set builder = NewBuilder(checks)
    builder.AddCondition vbNullString, "= $A$3"

    Assert.AreEqual CLng(0), builder.ConditionCount, _
                    "A criterion naming no variable adds nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestACriterionWithNoVariableIsIgnored", Err.Number, Err.Description
End Sub

'@sub-title Verify clearing empties the criteria set.
'@TestMethod("FormulaBuilder")
Public Sub TestClearEmptiesTheCriteriaSet()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestClearEmptiesTheCriteriaSet"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking

    Set builder = NewBuilder(checks)
    builder.AddCondition ROW_CHOICE_VARIABLE, "= $A$3"
    builder.ClearConditions

    Assert.AreEqual CLng(0), builder.ConditionCount, "Clearing leaves no criteria"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestClearEmptiesTheCriteriaSet", Err.Number, Err.Description
End Sub

'@sub-title Verify the first criterion after a build starts a new set.
'@details
'This is the discipline the builder was extracted to own. An arm that forgets to
'clear used to produce a formula carrying the criteria of the previous cell, and
'a formula with one criterion too many still reads as a formula, so no test could
'catch it.
'@TestMethod("FormulaBuilder")
Public Sub TestABuildClosesTheCriteriaSet()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestABuildClosesTheCriteriaSet"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim formObject As Formulas
    Dim firstFormula As String

    Set builder = NewBuilder(checks)
    Set formObject = Formulas.Create(dict, fData, COUNT_CALL_FUNCTION)

    builder.AddCondition ROW_CHOICE_VARIABLE, "= $A$3"
    firstFormula = builder.ExcelFormula(formObject)

    Assert.AreEqual CLng(1), builder.ConditionCount, _
                    "The criteria of the formula just built are still readable"

    builder.AddCondition ROW_CHOICE_VARIABLE, "= $A$4"

    Assert.AreEqual CLng(1), builder.ConditionCount, _
                    "The criterion after a build starts a new set, and the count is " & _
                    builder.ConditionCount
    Assert.IsTrue (Len(firstFormula) > 0), _
                  "The first formula should have been built, and it is [" & _
                  firstFormula & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestABuildClosesTheCriteriaSet", Err.Number, Err.Description
End Sub

'@sub-title Verify a condition that refuses a variable reaches the owner report.
'@details
'`FormulaCondition` files its validation failures into a store of its own, and
'until the builder pulled them nobody read that store: a table whose condition
'was dropped reached the delivered file with the analysis phase reading clean.
'@TestMethod("FormulaBuilder")
Public Sub TestARefusedConditionReachesTheOwnerReport()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestARefusedConditionReachesTheOwnerReport"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim formObject As Formulas
    Dim built As String

    Set builder = NewBuilder(checks)
    Set formObject = Formulas.Create(dict, fData, COUNT_CALL_FUNCTION)

    'A name the dictionary does not carry: the condition refuses it and says so.
    'ExcelFormula is a Property Get, so the answer is assigned. Calling it as a
    'statement is a compile error, and a compile error reaches the VBE as a
    'modal that stops the whole headless run.
    builder.AddCondition "no_such_variable_at_all", "= $A$3"
    built = builder.ExcelFormula(formObject)

    Assert.IsTrue checks.Length() > 0, _
                  "A condition that refused its variable should reach the owner report"
    Assert.IsTrue InStr(1, Messages(checks), "Validation failed") > 0, _
                  "The owner report should carry what the condition refused, and it holds " & _
                  Messages(checks)

    'The formula still comes back. ParsedCustomFormula builds the COUNTIFS from
    'the criterion text and reads no dictionary, so a name the dictionary does
    'not carry reaches the cell as a table column that is not there. The entry
    'above is the only warning the build gives, which is why it is pulled.
    Assert.IsTrue InStr(1, built, "no_such_variable_at_all") > 0, _
                  "The formula is built over the refused name, and it reads " & built

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARefusedConditionReachesTheOwnerReport", Err.Number, Err.Description
End Sub

'@sub-title Verify no formula comes back when there is no summary function.
'@TestMethod("FormulaBuilder")
Public Sub TestNoSummaryFunctionGivesNoFormula()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestNoSummaryFunctionGivesNoFormula"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking

    Set builder = NewBuilder(checks)
    builder.AddCondition ROW_CHOICE_VARIABLE, "= $A$3"

    Assert.AreEqual vbNullString, builder.ExcelFormula(Nothing), _
                    "A builder with no summary function answers with no formula"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestNoSummaryFunctionGivesNoFormula", Err.Number, Err.Description
End Sub

'@section The text helpers
'===============================================================================

'@sub-title Verify the empty string literal is two double quotes.
'@TestMethod("FormulaBuilder")
Public Sub TestEmptyTextIsTwoQuotes()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestEmptyTextIsTwoQuotes"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking

    Set builder = NewBuilder(checks)

    Assert.AreEqual Chr(34) & Chr(34), builder.EmptyText, _
                    "The empty string of a formula is two double quotes"
    Assert.AreEqual Chr(34) & "<>" & Chr(34), builder.NotEmptyText, _
                    "The not blank criteria of COUNTIFS is a quoted operator"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestEmptyTextIsTwoQuotes", Err.Number, Err.Description
End Sub

'@sub-title Verify a percentage guards its division.
'@TestMethod("FormulaBuilder")
Public Sub TestAPercentageGuardsItsDivision()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAPercentageGuardsItsDivision"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim formulaText As String

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    formulaText = builder.Percentage(sh.Cells(5, 2), sh.Cells(3, 1))

    Assert.IsTrue (InStr(1, formulaText, "ISERR") > 0), _
                  "A percentage answers blank on a zero denominator, and it reads [" & _
                  formulaText & "]"
    Assert.IsTrue (InStr(1, formulaText, "$B$5") > 0), _
                  "The denominator column is held fixed, and it reads [" & _
                  formulaText & "]"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAPercentageGuardsItsDivision", Err.Number, Err.Description
End Sub

'@sub-title Verify a guarded formula blanks when its guard cell is empty.
'@TestMethod("FormulaBuilder")
Public Sub TestAGuardedFormulaBlanksOnAnEmptyGuard()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAGuardedFormulaBlanksOnAnEmptyGuard"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim formulaText As String

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    formulaText = builder.Condition(sh.Cells(2, 3), "1+1")

    Assert.IsTrue (InStr(1, formulaText, "IF(") > 0), _
                  "The guard is an IF, and it reads [" & formulaText & "]"
    Assert.IsTrue (InStr(1, formulaText, "1+1") > 0), _
                  "The guarded formula is kept, and it reads [" & formulaText & "]"
    Assert.AreEqual vbNullString, builder.Condition(sh.Cells(2, 3), vbNullString), _
                    "There is nothing to guard when there is no formula"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAGuardedFormulaBlanksOnAnEmptyGuard", Err.Number, Err.Description
End Sub

'@section Writing formulas
'===============================================================================

'@sub-title Verify a formula reaches its cell and nothing is reported.
'@TestMethod("FormulaBuilder")
Public Sub TestAFormulaReachesItsCell()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAFormulaReachesItsCell"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim written As Boolean

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    written = builder.WriteFormula(sh.Cells(1, 1), "1+1")
    Application.Calculate

    Assert.IsTrue written, "A formula Excel accepts should report that it landed"
    Assert.AreEqual "=1+1", CStr(sh.Cells(1, 1).Formula), "The cell holds the formula"
    Assert.AreEqual CLng(2), CLng(sh.Cells(1, 1).Value), "And it computes"
    Assert.AreEqual CLng(0), checks.Length(), _
                    "A formula that landed should report nothing, and it reported " & _
                    Messages(checks)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFormulaReachesItsCell", Err.Number, Err.Description
End Sub

'@sub-title Verify an empty formula is reported and the cell is left alone.
'@TestMethod("FormulaBuilder")
Public Sub TestAnEmptyFormulaIsReported()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAnEmptyFormulaIsReported"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim written As Boolean

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    written = builder.WriteFormula(sh.Cells(2, 1), vbNullString)

    Assert.IsFalse written, "There is nothing to write when the text is empty"
    Assert.AreEqual CLng(1), checks.Length(), _
                    "The cell that has no formula should be named once, and the " & _
                    "report holds " & Messages(checks)
    Assert.IsTrue (InStr(1, Messages(checks), "No formula could be built") > 0), _
                  "The message says no formula could be built, and it reads " & _
                  Messages(checks)
    Assert.IsTrue (InStr(1, Messages(checks), CONTEXT_ID) > 0), _
                  "The message names the table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnEmptyFormulaIsReported", Err.Number, Err.Description
End Sub

'@sub-title Verify a formula Excel refuses is reported and the cell is cleared.
'@details
'Excel is the only judge the builder needs: it rejects a malformed formula at
'assignment time. This test writes a malformed formula. A semantically wrong one
'is accepted and answers with a reference error, which the builder leaves alone.
'@TestMethod("FormulaBuilder")
Public Sub TestARefusedFormulaIsReported()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestARefusedFormulaIsReported"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim written As Boolean

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    written = builder.WriteFormula(sh.Cells(3, 1), "SUM(")

    Assert.IsFalse written, "A formula Excel refuses should report that it did not land"
    Assert.AreEqual vbNullString, CStr(sh.Cells(3, 1).Formula), _
                    "A refused formula leaves the cell empty"
    Assert.IsTrue (InStr(1, Messages(checks), "refused") > 0), _
                  "The message says Excel refused it, and the report holds " & _
                  Messages(checks)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestARefusedFormulaIsReported", Err.Number, Err.Description
End Sub

'@sub-title Verify a formula over 255 characters reaches its cell and computes.
'@details
'Range.FormulaArray refuses the assignment of a longer formula, so it goes in
'through a stub and a replace. The answer it computes is the assertion worth
'writing: reading the text back says nothing about whether the entry was right.
'@TestMethod("FormulaBuilder")
Public Sub TestALongFormulaReachesItsCell()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestALongFormulaReachesItsCell"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim formulaText As String
    Dim written As Boolean
    Dim onesCount As Long

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    formulaText = LongFormulaText()
    onesCount = (Len(formulaText) + 1) / 2
    written = builder.WriteFormula(sh.Cells(4, 1), formulaText)
    Application.Calculate

    Assert.IsTrue (Len(formulaText) > FORMULA_ARRAY_LIMIT), _
                  "The fixture formula is longer than an array formula may be"
    Assert.IsTrue written, "A long formula should reach its cell"
    Assert.AreEqual CLng(onesCount), CLng(sh.Cells(4, 1).Value), _
                    "And it computes the same answer, which is " & onesCount

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALongFormulaReachesItsCell", Err.Number, Err.Description
End Sub

'@sub-title Verify a formula is copied down and a single cell is left alone.
'@TestMethod("FormulaBuilder")
Public Sub TestAFormulaIsCopiedDown()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAFormulaIsCopiedDown"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    builder.WriteFormula sh.Cells(1, 5), "1+1", False
    builder.FillDown sh.Cells(1, 5), sh.Range(sh.Cells(1, 5), sh.Cells(3, 5))

    Assert.IsTrue (InStr(1, CStr(sh.Cells(3, 5).Formula), "1+1") > 0), _
                  "The last cell of the range carries the formula, and it holds [" & _
                  CStr(sh.Cells(3, 5).Formula) & "]"

    builder.FillDown sh.Cells(1, 5), sh.Cells(1, 5)

    Assert.AreEqual CLng(0), checks.Length(), _
                    "A destination of one cell is left alone with nothing reported, " & _
                    "and the report holds " & Messages(checks)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFormulaIsCopiedDown", Err.Number, Err.Description
End Sub

'@section Reporting
'===============================================================================

'@sub-title Verify a key names the owner and the table it wrote for.
'@details
'AnalysisOutput pours the entries of several classes over several tables into one
'report, and Checking.Add raises on a duplicate key. A bare counter collided with
'the first entry of every other writer of the same table.
'@TestMethod("FormulaBuilder")
Public Sub TestAKeyNamesTheOwnerAndTheTable()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAKeyNamesTheOwnerAndTheTable"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim keyList As BetterArray
    Dim firstKey As String

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    builder.WriteFormula sh.Cells(6, 1), vbNullString
    builder.WriteFormula sh.Cells(7, 1), vbNullString

    Set keyList = checks.ListOfKeys
    firstKey = CStr(keyList.Item(keyList.LowerBound))

    Assert.AreEqual CLng(2), checks.Length(), _
                    "Two entries under two keys, and the report holds " & Messages(checks)
    Assert.AreEqual OWNER_NAME & "-" & CONTEXT_ID & "-F1", firstKey, _
                    "The key names the owner, the table and the entry"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAKeyNamesTheOwnerAndTheTable", Err.Number, Err.Description
End Sub

'@sub-title Verify a format keyword reaches the cell as an Excel number format.
'@TestMethod("FormulaBuilder")
Public Sub TestAFormatKeywordIsApplied()
    CustomTestSetTitles Assert, "FormulaBuilder", "TestAFormatKeywordIsApplied"
    On Error GoTo TestFail

    Dim builder As FormulaBuilder
    Dim checks As Checking
    Dim sh As Worksheet
    Dim beforeFormat As String

    Set builder = NewBuilder(checks)
    Set sh = ScratchSheet()

    beforeFormat = CStr(sh.Cells(9, 1).NumberFormat)
    builder.ApplyFormat "integer", sh.Cells(9, 1)

    Assert.IsTrue (CStr(sh.Cells(9, 1).NumberFormat) <> beforeFormat), _
                  "The integer keyword changes the number format, and it reads [" & _
                  CStr(sh.Cells(9, 1).NumberFormat) & "]"

    builder.ApplyFormat vbNullString, sh.Cells(10, 1)

    Assert.AreEqual CLng(0), checks.Length(), _
                    "Formatting reports nothing, and the report holds " & Messages(checks)

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAFormatKeywordIsApplied", Err.Number, Err.Description
End Sub
