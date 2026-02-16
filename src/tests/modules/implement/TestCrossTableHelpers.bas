Attribute VB_Name = "TestCrossTableHelpers"

Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Implement")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private FixtureWorkbook As Workbook
Private SpecsStub As GraphTablesSpecsStub

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set FixtureWorkbook = TestHelpers.NewWorkbook
    Set SpecsStub = New GraphTablesSpecsStub
    SpecsStub.Configure ScopeBivariate, "TEST_TABLE"
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not FixtureWorkbook Is Nothing Then
        TestHelpers.DeleteWorkbook FixtureWorkbook
        Set FixtureWorkbook = Nothing
    End If
    Set SpecsStub = Nothing
    TestHelpers.RestoreApp
End Sub

'@section Layout Planner
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestLayoutPlannerRequiresSpecifications()
    On Error GoTo ExpectError
        CrossTableLayoutPlanner.Create Nothing, FixtureWorkbook.Worksheets(1)
        Assert.Fail "Planner should reject missing specifications"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, "Planner must validate specs argument"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLayoutPlannerStoresDependencies()
    Dim planner As CrossTableLayoutPlanner
    Set planner = CrossTableLayoutPlanner.Create(SpecsStub, FixtureWorkbook.Worksheets(1))

    Assert.IsTrue planner.Specifications Is SpecsStub, "Planner should expose supplied specs"
    Assert.IsTrue planner.TargetSheet Is FixtureWorkbook.Worksheets(1), "Planner should expose supplied sheet"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLayoutPlannerComputesStartCoordinates()
    Dim planner As CrossTableLayoutPlanner
    Dim summary As CrossTableLayoutSummary

    SpecsStub.SetIsNewSection True
    SpecsStub.SetRowCategories "Male", "Female"
    SpecsStub.SetColumnCategories "Yes", "No"
    SpecsStub.SetHasTotal True
    SpecsStub.SetHasMissing True

    Set planner = CrossTableLayoutPlanner.Create(SpecsStub, FixtureWorkbook.Worksheets(1))
    Set summary = planner.ComputeLayout()

    Assert.AreEqual 10, summary.StartRow, "Bivariate tables should reserve spacing before data"
    Assert.AreEqual 3, summary.StartColumn, "Start column should default to C"
    Assert.AreEqual 2, summary.HeaderRowCount, "Header rows should reflect table structure"
    Assert.AreEqual 4, summary.DataRowCount, "Data rows should include totals and missing rows"
    Assert.AreEqual 3, summary.DataColumnCount, "Data columns should include totals"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLayoutPlannerRespectsExistingNamedRanges()
    Dim sheet As Worksheet
    Dim planner As CrossTableLayoutPlanner
    Dim summary As CrossTableLayoutSummary

    Set sheet = FixtureWorkbook.Worksheets(1)
    FixtureWorkbook.Names.Add Name:="STARTROW_TEST_TABLE", _
        RefersTo:="=" & sheet.Cells(42, 3).Address(True, True, xlA1, True)
    FixtureWorkbook.Names.Add Name:="STARTCOL_TEST_TABLE", _
        RefersTo:="=" & sheet.Cells(41, 7).Address(True, True, xlA1, True)

    Set planner = CrossTableLayoutPlanner.Create(SpecsStub, sheet)
    Set summary = planner.ComputeLayout()

    Assert.AreEqual 42, summary.StartRow, "Existing names should drive start row"
    Assert.AreEqual 6, summary.StartColumn, "Named column should map back to start column"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLayoutPlannerUsesPreviousForTimeSeries()
    Dim sheet As Worksheet
    Dim previousSpecs As GraphTablesSpecsStub
    Dim planner As CrossTableLayoutPlanner
    Dim summary As CrossTableLayoutSummary

    Set sheet = FixtureWorkbook.Worksheets(1)

    Set previousSpecs = New GraphTablesSpecsStub
    previousSpecs.Configure ScopeTimeSeries, "PREV_TABLE"
    previousSpecs.SetIsNewSection True
    FixtureWorkbook.Names.Add Name:="STARTROW_PREV_TABLE", _
        RefersTo:="=" & sheet.Cells(30, 3).Address(True, True, xlA1, True)

    SpecsStub.Configure ScopeTimeSeries, "CURR_TABLE"
    SpecsStub.SetIsNewSection False
    SpecsStub.SetPrevious previousSpecs

    Set planner = CrossTableLayoutPlanner.Create(SpecsStub, sheet)
    Set summary = planner.ComputeLayout()

    Assert.AreEqual 30, summary.StartRow, "Time series tables should reuse previous start row"
End Sub

'@section Range Registrar
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestRangeRegistrarRequiresWorksheet()
    On Error GoTo ExpectError
        CrossTableRangeRegistrar.Create Nothing
        Assert.Fail "Registrar should reject missing worksheet"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ObjectNotInitialized, Err.Number, "Registrar must validate worksheet argument"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestRangeRegistrarRegistersNames()
    Dim registrar As CrossTableRangeRegistrar
    Dim summary As CrossTableLayoutSummary
    Dim sheet As Worksheet
    Dim refersTo As String

    Set sheet = FixtureWorkbook.Worksheets(1)
    Set registrar = CrossTableRangeRegistrar.Create(sheet)
    Set summary = New CrossTableLayoutSummary
    summary.Initialise 8, 3, 2, 4, 3

    registrar.RegisterLayoutName "STARTROW_TEST", summary
    registrar.RegisterLayoutName "TABLE_TEST", summary

    refersTo = FixtureWorkbook.Names("STARTROW_TEST").RefersTo
    Assert.IsTrue InStr(1, refersTo, "!$C$8", vbTextCompare) > 0, "STARTROW should target start cell"

    refersTo = FixtureWorkbook.Names("TABLE_TEST").RefersTo
    Assert.IsTrue InStr(1, refersTo, "!$C$8", vbTextCompare) > 0, "Table range should start at computed coordinates"

    registrar.RemoveName "STARTROW_TEST"
    Assert.IsFalse TestHelpers.NamedRangeExists("STARTROW_TEST", FixtureWorkbook), _
                 "RemoveName should delete workbook-level name"
End Sub

'@section Write Buffer
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestWriteBufferCommitsMatrices()
    Dim buffer As CrossTableWriteBuffer
    Dim sheet As Worksheet
    Dim header As Variant
    Dim body As Variant

    Set sheet = FixtureWorkbook.Worksheets(1)
    header = TestHelpers.RowsToMatrix(Array(Array("H1", "H2")))
    body = TestHelpers.RowsToMatrix(Array(Array(1, 2), Array(3, 4)))

    Set buffer = CrossTableServices.BuildWriteBuffer(5, 3)
    buffer.AssignHeader header
    buffer.AssignBody body
    buffer.Commit sheet

    Assert.AreEqual "H1", sheet.Cells(5, 3).Value, "Header should be written at start cell"
    Assert.AreEqual 3, sheet.Cells(7, 3).Value, "Body should start after header rows"
    Assert.AreEqual 4, sheet.Cells(8, 4).Value, "Body values should be preserved"
End Sub

'@section Condition Set
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestConditionSetInitialises()
    Dim conditions As CrossTableConditionSet
    Set conditions = New CrossTableConditionSet
    conditions.Initialise

    Assert.AreEqual 0, conditions.Count, "Initialised condition set should be empty"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestConditionSetManagesDuplicates()
    Dim conditions As CrossTableConditionSet
    Set conditions = New CrossTableConditionSet
    conditions.Initialise

    conditions.Add "age", "age>0"
    conditions.Add "AGE", "age>=0"
    conditions.Add "sex", "sex<>" & Chr$(34) & Chr$(34)

    Assert.AreEqual 2, conditions.Count, "Duplicate variables should be merged"
    Assert.AreEqual "age", conditions.VariableAt(1), "Variable order should be stable"
    Assert.AreEqual "age>=0", conditions.TestAt(1), "Latest condition should replace duplicate"
End Sub

'@section Formula Builder
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestFormulaBuilderPercentage()
    Dim builder As CrossTableFormulaBuilder
    Dim formulaText As String
    Dim sheet As Worksheet

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1").Value = 10
    sheet.Range("B1").Value = 20

    Set builder = CrossTableFormulaBuilder.Create(sheet)
    formulaText = builder.BuildPercentage(sheet.Range("A1"), sheet.Range("B1"))

    Assert.AreEqual "IF(ISERR(A1/B$1),"""""",A1/B$1)", formulaText, _
                     "Percentage formula should wrap potential division errors"
    Assert.IsTrue builder.Validate(formulaText), "Formula should evaluate successfully"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestFormulaBuilderConditional()
    Dim builder As CrossTableFormulaBuilder
    Dim sheet As Worksheet
    Dim formulaText As String

    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("C1").Value = "value"

    Set builder = CrossTableFormulaBuilder.Create(sheet)
    formulaText = builder.BuildConditional(sheet.Range("C1"), "VALUE")

    Assert.AreEqual "IF(C1="""""","""""",VALUE)", formulaText, _
                     "Conditional formula should guard against empty cells"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestFormulaBuilderValidateDetectsInvalid()
    Dim builder As CrossTableFormulaBuilder
    Dim sheet As Worksheet

    Set sheet = FixtureWorkbook.Worksheets(1)
    Set builder = CrossTableFormulaBuilder.Create(sheet)

    Assert.IsFalse builder.Validate("SUM("), "Malformed formulas should fail validation"
End Sub

'@section Legacy Adapter
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestLegacyAdapterPlanAndRegister()
    Dim adapter As CrossTableLegacyAdapter
    Dim summary As CrossTableLayoutSummary

    SpecsStub.SetIsNewSection True
    SpecsStub.SetRowCategories "Cat1"
    SpecsStub.SetColumnCategories "Col1"

    Set adapter = CrossTableLegacyAdapter.Create(SpecsStub, FixtureWorkbook.Worksheets(1))
    Set summary = adapter.PlanAndRegister("LEG")

    Assert.ObjectExists summary, "CrossTableLayoutSummary", "Adapter should return a layout summary"
    Assert.IsTrue TestHelpers.NamedRangeExists("LEG_STARTROW", FixtureWorkbook), _
                 "Adapter should register the start row name"
    Assert.IsTrue TestHelpers.NamedRangeExists("LEG_TABLE", FixtureWorkbook), _
                 "Adapter should register the table range name"
    Assert.IsTrue adapter.PlanDurationMilliseconds >= 0, _
                 "Plan duration should be captured"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLegacyAdapterWriteContent()
    Dim adapter As CrossTableLegacyAdapter
    Dim header As Variant
    Dim body As Variant
    Dim sheet As Worksheet

    SpecsStub.SetIsNewSection True
    SpecsStub.SetRowCategories "Cat1"
    SpecsStub.SetColumnCategories "Col1"

    Set sheet = FixtureWorkbook.Worksheets(1)
    Set adapter = CrossTableLegacyAdapter.Create(SpecsStub, sheet)
    adapter.PlanAndRegister "LEG"

    header = TestHelpers.RowsToMatrix(Array(Array("Heading")))
    body = TestHelpers.RowsToMatrix(Array(Array("Value")))

    adapter.WriteContent header, body

    Assert.AreEqual "Heading", sheet.Cells(adapter.LayoutSummary.StartRow, _
                                           adapter.LayoutSummary.StartColumn).Value, _
                     "Adapter should write header values"
    Assert.AreEqual "Value", sheet.Cells(adapter.LayoutSummary.StartRow + 1, _
                                         adapter.LayoutSummary.StartColumn).Value, _
                     "Adapter should write body values"
    Assert.IsTrue adapter.WriteDurationMilliseconds >= 0, _
                 "Write duration should be captured"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLegacyAdapterBuildPercentageFormula()
    Dim adapter As CrossTableLegacyAdapter
    Dim sheet As Worksheet
    Dim formulaText As String

    SpecsStub.SetIsNewSection True
    Set sheet = FixtureWorkbook.Worksheets(1)
    sheet.Range("A1").Value = 1
    sheet.Range("B1").Value = 2

    Set adapter = CrossTableLegacyAdapter.Create(SpecsStub, sheet)
    formulaText = adapter.BuildPercentageFormula(sheet.Range("A1"), sheet.Range("B1"))

    Assert.AreEqual "IF(ISERR(A1/B$1),"""""",A1/B$1)", formulaText, _
                     "Adapter should delegate to the formula builder"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLegacyAdapterConditionSet()
    Dim adapter As CrossTableLegacyAdapter
    Dim conditionSet As CrossTableConditionSet

    SpecsStub.SetIsNewSection True
    Set adapter = CrossTableLegacyAdapter.Create(SpecsStub, FixtureWorkbook.Worksheets(1))
    Set conditionSet = adapter.NewConditionSet

    Assert.AreEqual 0, conditionSet.Count, "Condition set should be initialised empty"
    conditionSet.Add "flag", "flag<>" & Chr$(34) & Chr$(34)
    Assert.AreEqual 1, conditionSet.Count, "Condition set should accept additions"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestLegacyAdapterWriteContentRequiresPlanning()
    Dim adapter As CrossTableLegacyAdapter

    SpecsStub.SetIsNewSection True
    Set adapter = CrossTableLegacyAdapter.Create(SpecsStub, FixtureWorkbook.Worksheets(1))

    On Error GoTo ExpectError
        adapter.WriteContent Array(), Array()
        Assert.Fail "WriteContent should require prior planning"
        Exit Sub
ExpectError:
    Assert.AreEqual ProjectError.ErrorUnexpectedState, Err.Number, _
                     "Adapter must enforce planning before writing"
End Sub

'@section Performance Tracker
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestPerformanceTrackerRecordsDuration()
    Dim tracker As CrossTablePerformanceTracker

    Set tracker = New CrossTablePerformanceTracker
    tracker.BeginMeasurement
    tracker.EndMeasurement

    Assert.IsTrue tracker.DurationMilliseconds >= 0, _
                 "Duration should be non-negative even for immediate measurements"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestPerformanceTrackerGuardsState()
    Dim tracker As CrossTablePerformanceTracker

    Set tracker = New CrossTablePerformanceTracker

    On Error GoTo ExpectErrorEnd
        tracker.EndMeasurement
        Assert.Fail "Ending without start should raise"
        Exit Sub
ExpectErrorEnd:
    Assert.AreEqual ProjectError.ErrorUnexpectedState, Err.Number, _
                     "Tracker should enforce valid state transitions"

    tracker.BeginMeasurement

    On Error GoTo ExpectErrorSecondStart
        tracker.BeginMeasurement
        Assert.Fail "Second start should raise"
        Exit Sub
ExpectErrorSecondStart:
    Assert.AreEqual ProjectError.ErrorUnexpectedState, Err.Number, _
                     "Tracker should reject overlapping measurements"

    tracker.EndMeasurement
End Sub

'@section Service Facade
'===============================================================================

'@TestMethod("CrossTableHelpers")
Private Sub TestServicesFactory()
    Dim services As CrossTableServices
    Dim planner As CrossTableLayoutPlanner
    Dim registrar As CrossTableRangeRegistrar

    Set services = New CrossTableServices
    Set planner = services.BuildLayoutPlanner(SpecsStub, FixtureWorkbook.Worksheets(1))
    Set registrar = services.BuildRangeRegistrar(FixtureWorkbook.Worksheets(1))

    Assert.ObjectExists planner, "CrossTableLayoutPlanner", "Services should return a layout planner"
    Assert.ObjectExists registrar, "CrossTableRangeRegistrar", "Services should return a range registrar"
End Sub

'@TestMethod("CrossTableHelpers")
Private Sub TestServicesBuildWriteBuffer()
    Dim services As CrossTableServices
    Dim buffer As CrossTableWriteBuffer

    Set services = New CrossTableServices
    Set buffer = services.BuildWriteBuffer(4, 2)

    Assert.ObjectExists buffer, "CrossTableWriteBuffer", "Facade should supply write buffer instances"
End Sub
