Attribute VB_Name = "TestAnalysisTablePolicyResolver"
Attribute VB_Description = "Unit tests covering AnalysisTablePolicyResolver behaviour"

Option Explicit
Option Private Module

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Unit tests covering AnalysisTablePolicyResolver behaviour")

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const SPEC_SHEET As String = "ResolverSpecs"
Private Const SPEC_TABLE_NAME As String = "T_ResolverSpecs"

Private Assert As ICustomTest
Private SpecSheet As Worksheet
Private SpecTable As ListObject
Private Enumerator As IAnalysisTableEnumerator
Private BuilderStub As AnalysisTableEnumeratorBuilderStub
Private Linelist As TableSpecsLinelistStub
Private DictionaryStub As AnalysisDictionaryStub
Private VariablesStub As AnalysisVariablesStub
Private VariablesCacheStub As AnalysisVariablesCacheStub
Private PolicyProviderStub As AnalysisTablePolicyProviderStub
Private IterationItems As BetterArray

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestAnalysisTablePolicyResolver"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    DeleteWorksheet SPEC_SHEET
    RestoreApp

    Set Assert = Nothing
    Set SpecSheet = Nothing
    Set SpecTable = Nothing
    Set Enumerator = Nothing
    Set BuilderStub = Nothing
    Set Linelist = Nothing
    Set DictionaryStub = Nothing
    Set VariablesStub = Nothing
    Set VariablesCacheStub = Nothing
    Set PolicyProviderStub = Nothing
    Set IterationItems = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    BusyApp
    Set SpecSheet = EnsureWorksheet(SPEC_SHEET)
    ClearWorksheet SpecSheet

    Set BuilderStub = New AnalysisTableEnumeratorBuilderStub
    Set Enumerator = AnalysisTableEnumerator.Create(BuilderStub.Self)

    Set Linelist = New TableSpecsLinelistStub
    Set DictionaryStub = New AnalysisDictionaryStub
    DictionaryStub.AddVariable "age"
    Linelist.SetDictionary DictionaryStub

    Set VariablesStub = New AnalysisVariablesStub
    VariablesStub.AddVariable "age", "choice_manual", "numeric"

    Set VariablesCacheStub = New AnalysisVariablesCacheStub
    VariablesCacheStub.Configure DictionaryStub, VariablesStub
    VariablesCacheStub.LinelistSpecificationsStub = Linelist

    Set PolicyProviderStub = New AnalysisTablePolicyProviderStub

    BuildSpecificationTable SectionArray("Section A", "Section A")
    Set IterationItems = Enumerator.Enumerate(SpecTable, Linelist, TypeUnivariate)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If

    If Not SpecSheet Is Nothing Then
        ClearWorksheet SpecSheet
    End If

    Set SpecTable = Nothing
    Set IterationItems = Nothing
End Sub

'@section Helpers
'===============================================================================
Private Sub BuildSpecificationTable(ByVal sections As Variant)
    Dim rowCount As Long
    Dim idx As Long
    Dim rows() As Variant
    Dim matrix As Variant
    Dim tableRange As Range

    rowCount = UBound(sections) - LBound(sections) + 1
    ReDim rows(0 To rowCount)
    rows(0) = Array("section", "table_id", "row", "percentage")

    For idx = 0 To rowCount - 1
        rows(idx + 1) = Array( _
            CStr(sections(LBound(sections) + idx)), _
            "table_" & CStr(idx + 1), _
            "age", _
            "yes" _
        )
    Next idx

    matrix = RowsToMatrix(rows)
    WriteMatrix SpecSheet.Range("A1"), matrix

    Set tableRange = SpecSheet.Range("A1").Resize(UBound(matrix, 1), UBound(matrix, 2))
    Set SpecTable = SpecSheet.ListObjects.Add(SourceType:=xlSrcRange, Source:=tableRange, XlListObjectHasHeaders:=xlYes)
    SpecTable.Name = SPEC_TABLE_NAME
End Sub

Private Function SectionArray(ParamArray sectionValues() As Variant) As Variant
    SectionArray = sectionValues
End Function

Private Function FirstIterationItem() As IAnalysisTableIterationItem
    Set FirstIterationItem = IterationItems.Item(1)
End Function

'@section Tests
'===============================================================================
'@TestMethod("AnalysisTablePolicyResolver")
Public Sub TestResolveReturnsPolicyFlags()
    CustomTestSetTitles Assert, "AnalysisTablePolicyResolver", "TestResolveReturnsPolicyFlags"
    Dim policyStub As AnalysisTypePolicyStub
    Dim resolver As IAnalysisTablePolicyResolver
    Dim result As IAnalysisTablePolicyResult

    On Error GoTo Fail

    Set policyStub = New AnalysisTypePolicyStub
    policyStub.Configure TypeUnivariate, True, True, False, True, False
    PolicyProviderStub.AddPolicy TypeUnivariate, policyStub

    Set resolver = AnalysisTablePolicyResolver.Create(Linelist, VariablesCacheStub, PolicyProviderStub)
    Set result = resolver.Resolve(FirstIterationItem())

    Assert.IsTrue result.IsValid, "Policy should mark specification as valid"
    Assert.IsTrue result.HasPercent, "HasPercent flag should mirror policy"
    Assert.IsFalse result.HasTotal, "HasTotal flag should mirror policy"
    Assert.IsTrue result.HasMissing, "HasMissing flag should mirror policy"
    Assert.IsFalse result.HasGraph, "HasGraph flag should mirror policy"
    Assert.AreSameObj policyStub, result.Policy, "Resolved policy should match stub"
    Assert.IsTrue (Not result.Context Is Nothing), "Policy context should be available"
    Assert.AreSameObj FirstIterationItem(), result.IterationItem, "Iteration item should be preserved"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestResolveReturnsPolicyFlags"
End Sub

'@TestMethod("AnalysisTablePolicyResolver")
Public Sub TestInvalidateDelegatesToCache()
    CustomTestSetTitles Assert, "AnalysisTablePolicyResolver", "TestInvalidateDelegatesToCache"
    Dim resolver As IAnalysisTablePolicyResolver

    On Error GoTo Fail

    Set resolver = AnalysisTablePolicyResolver.Create(Linelist, VariablesCacheStub, PolicyProviderStub)
    resolver.Invalidate

    Assert.IsTrue VariablesCacheStub.InvalidateCalled, "Invalidate should propagate to variables cache"

    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestInvalidateDelegatesToCache"
End Sub
