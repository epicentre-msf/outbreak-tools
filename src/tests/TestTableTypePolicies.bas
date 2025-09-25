Attribute VB_Name = "TestTableTypePolicies"
Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests covering TableSpecs table-type policies and factory")

Private Assert As Object
Private Factory As TableTypePolicyFactory

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Factory = New TableTypePolicyFactory
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Factory = Nothing
    Set Assert = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Function NewContext(ByVal tableType As Byte, _
                            ByVal rowVariable As String, _
                            ByVal columnVariable As String) As FakeTableSpecsPolicyContext
    Set NewContext = New FakeTableSpecsPolicyContext
    NewContext.Configure tableType, rowVariable, columnVariable
End Function

'@section Tests
'===============================================================================

'@TestMethod("TableTypePolicies")
Private Sub TestFactoryCachesPolicyInstances()
    Dim firstPolicy As ITableTypePolicy
    Dim secondPolicy As ITableTypePolicy

    On Error GoTo Fail

    Set firstPolicy = Factory.Create(TABLE_TYPE_UNIVARIATE)
    Set secondPolicy = Factory.Create(TABLE_TYPE_UNIVARIATE)

    Assert.AreSameObj firstPolicy, secondPolicy, _
        "Factory should cache and reuse policy instances"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestFactoryCachesPolicyInstances"
End Sub

'@TestMethod("TableTypePolicies")
Private Sub TestGlobalSummaryPolicyValidation()
    Dim policy As ITableTypePolicy
    Dim context As FakeTableSpecsPolicyContext

    On Error GoTo Fail

    Set policy = Factory.Create(TABLE_TYPE_GLOBAL_SUMMARY)
    Set context = NewContext(TABLE_TYPE_GLOBAL_SUMMARY, "", "")

    context.SetColumnValue "label", "Cases"
    context.SetColumnValue "function", "sum"

    Assert.IsFalse policy.HasPercent(context), "Global summary should never expose percentages"
    Assert.IsFalse policy.HasTotal(context), "Global summary should never expose totals"
    Assert.IsFalse policy.HasMissing(context), "Global summary should never expose missing values"
    Assert.IsFalse policy.HasGraph(context), "Global summary should never expose graphs"
    Assert.IsTrue policy.IsValid(context), "Label and function presence should mark the table valid"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestGlobalSummaryPolicyValidation"
End Sub

'@TestMethod("TableTypePolicies")
Private Sub TestUnivariatePolicyRules()
    Dim policy As ITableTypePolicy
    Dim context As FakeTableSpecsPolicyContext

    On Error GoTo Fail

    Set policy = Factory.Create(TABLE_TYPE_UNIVARIATE)
    Set context = NewContext(TABLE_TYPE_UNIVARIATE, "sex", "")

    context.SetColumnValue "percentage", "yes"
    context.SetColumnValue "missing", "yes"
    context.SetColumnValue "graph", "yes"

    context.SetVariablePresence "sex", True
    context.SetVariableControl "sex", "choice_manual"

    Assert.IsTrue policy.HasPercent(context), "Percentage flag should be honoured"
    Assert.IsTrue policy.HasTotal(context), "Univariate tables always expose totals"
    Assert.IsTrue policy.HasMissing(context), "Missing flag should be honoured"
    Assert.IsTrue policy.HasGraph(context), "Graph flag should be honoured"
    Assert.IsTrue policy.IsValid(context), "Choice controls should validate"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestUnivariatePolicyRules"
End Sub

'@TestMethod("TableTypePolicies")
Private Sub TestBivariatePolicyValidation()
    Dim policy As ITableTypePolicy
    Dim context As FakeTableSpecsPolicyContext

    On Error GoTo Fail

    Set policy = Factory.Create(TABLE_TYPE_BIVARIATE)
    Set context = NewContext(TABLE_TYPE_BIVARIATE, "sex", "age_group")

    context.SetColumnValue "percentage", "row"
    context.SetColumnValue "missing", "all"
    context.SetColumnValue "graph", "both"

    context.SetVariablePresence "sex", True
    context.SetVariableControl "sex", "choice_formula"

    context.SetVariablePresence "age_group", True
    context.SetVariableControl "age_group", "choice_manual"

    Assert.IsTrue policy.HasPercent(context), "Row percentages should be recognised"
    Assert.IsTrue policy.HasTotal(context), "Bivariate tables expose totals"
    Assert.IsTrue policy.HasMissing(context), "Missing 'all' selection should be recognised"
    Assert.IsTrue policy.HasGraph(context), "Graph selection should be honoured"
    Assert.IsTrue policy.IsValid(context), "Both controls configured as choices should validate"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestBivariatePolicyValidation"
End Sub

'@TestMethod("TableTypePolicies")
Private Sub TestTimeSeriesPolicyRules()
    Dim policy As ITableTypePolicy
    Dim context As FakeTableSpecsPolicyContext

    On Error GoTo Fail

    Set policy = Factory.Create(TABLE_TYPE_TIME_SERIES)
    Set context = NewContext(TABLE_TYPE_TIME_SERIES, "onset_date", "outcome")

    context.SetColumnValue "percentage", "row"
    context.SetColumnValue "total", "yes"
    context.SetColumnValue "column", "outcome"
    context.SetColumnValue "missing", "yes"

    context.SetVariableType "onset_date", "date"
    context.SetVariableControl "outcome", "choice_manual"

    Assert.IsTrue policy.HasPercent(context), "Row percentages should require totals"
    Assert.IsTrue policy.HasTotal(context), "Total flag with column should validate"
    Assert.IsTrue policy.HasMissing(context), "Missing flag should respect the column"
    Assert.IsFalse policy.HasGraph(context), "Time series should not expose graphs"
    Assert.IsTrue policy.IsValid(context), "Date row variable and choice column should validate"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestTimeSeriesPolicyRules"
End Sub

'@TestMethod("TableTypePolicies")
Private Sub TestSpatialPolicyValidation()
    Dim policy As ITableTypePolicy
    Dim context As FakeTableSpecsPolicyContext

    On Error GoTo Fail

    Set policy = Factory.Create(TABLE_TYPE_SPATIAL)
    Set context = NewContext(TABLE_TYPE_SPATIAL, "facility", "outcome")

    context.SetColumnValue "percentage", "yes"
    context.SetColumnValue "column", "outcome"
    context.SetColumnValue "missing", "yes"
    context.SetColumnValue "graph", "yes"

    context.AddDictionaryVariable "adm1_facility"
    context.SetVariableControl "outcome", "choice_formula"

    Assert.IsTrue policy.HasPercent(context), "Percentage flag should rely on totals"
    Assert.IsTrue policy.HasTotal(context), "Column presence should enable totals"
    Assert.IsTrue policy.HasMissing(context), "Missing flag should respect the column"
    Assert.IsTrue policy.HasGraph(context), "Graph flag should be honoured"
    Assert.IsTrue policy.IsValid(context), "Spatial prefix and choice column should validate"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSpatialPolicyValidation"
End Sub

'@TestMethod("TableTypePolicies")
Private Sub TestSpatioTemporalPolicyValidation()
    Dim policy As ITableTypePolicy
    Dim context As FakeTableSpecsPolicyContext

    On Error GoTo Fail

    Set policy = Factory.Create(TABLE_TYPE_SPATIO_TEMPORAL)
    Set context = NewContext(TABLE_TYPE_SPATIO_TEMPORAL, "onset_date", "facility")

    context.SetColumnValue "graph", "yes"
    context.SetVariableType "onset_date", "date"
    context.AddDictionaryVariable "hf_facility"

    Assert.IsFalse policy.HasPercent(context), "Spatio-temporal tables should not expose percentages"
    Assert.IsFalse policy.HasTotal(context), "Spatio-temporal tables should not expose totals"
    Assert.IsFalse policy.HasMissing(context), "Spatio-temporal tables should not expose missing values"
    Assert.IsTrue policy.HasGraph(context), "Graph flag should be honoured"
    Assert.IsTrue policy.IsValid(context), "Date row variable with spatial column should validate"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestSpatioTemporalPolicyValidation"
End Sub

