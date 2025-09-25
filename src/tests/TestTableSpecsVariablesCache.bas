Attribute VB_Name = "TestTableSpecsVariablesCache"
Option Explicit
Option Private Module

'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Tests for the TableSpecsVariablesCache helper")

Private Const PRIMARY_SHEET As String = "TableSpecsCachePrimary"
Private Const SECONDARY_SHEET As String = "TableSpecsCacheSecondary"

Private Assert As Object
Private LinelistStub As TableSpecsLinelistStub
Private Cache As ITableSpecsVariablesCache
Private PrimaryDictionary As ILLdictionary
Private SecondaryDictionary As ILLdictionary

'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    DeleteWorksheet PRIMARY_SHEET
    DeleteWorksheet SECONDARY_SHEET
    Set SecondaryDictionary = Nothing
    Set PrimaryDictionary = Nothing
    Set Cache = Nothing
    Set LinelistStub = Nothing
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================

'@TestInitialize
Private Sub TestInitialize()
    PrepareDictionaryFixture PRIMARY_SHEET
    PrepareDictionaryFixture SECONDARY_SHEET

    Set LinelistStub = New TableSpecsLinelistStub
    Set PrimaryDictionary = LLdictionary.Create(ThisWorkbook.Worksheets(PRIMARY_SHEET), 1, 1)
    LinelistStub.SetDictionary PrimaryDictionary
    Set Cache = TableSpecsVariablesCache.Create(LinelistStub)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Cache = Nothing
    Set LinelistStub = Nothing
    Set SecondaryDictionary = Nothing
    Set PrimaryDictionary = Nothing
End Sub

'@section Tests
'===============================================================================

'@TestMethod("TableSpecsVariablesCache")
Private Sub TestVariablesAccessorCachesInstance()
    Dim firstVariables As ILLVariables
    Dim secondVariables As ILLVariables

    On Error GoTo Fail

    Set firstVariables = Cache.Variables
    Set secondVariables = Cache.Variables

    Assert.AreSameObj firstVariables, secondVariables, "Variables accessor should reuse the cached instance"
    Assert.AreSameObj Cache.Dictionary, PrimaryDictionary, "Cache should expose the linelist dictionary"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestVariablesAccessorCachesInstance"
End Sub

'@TestMethod("TableSpecsVariablesCache")
Private Sub TestInvalidateClearsCachedReferences()
    Dim firstVariables As ILLVariables
    Dim secondVariables As ILLVariables

    On Error GoTo Fail

    Set firstVariables = Cache.Variables
    Cache.Invalidate
    Set secondVariables = Cache.Variables

    Assert.IsFalse firstVariables Is secondVariables, "Invalidate should force the cache to rebuild variables"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestInvalidateClearsCachedReferences"
End Sub

'@TestMethod("TableSpecsVariablesCache")
Private Sub TestRefreshRehydratesDictionaryChanges()
    Dim firstDictionary As ILLdictionary
    Dim secondDictionary As ILLdictionary
    Dim firstVariables As ILLVariables
    Dim refreshedVariables As ILLVariables

    On Error GoTo Fail

    Set firstDictionary = Cache.Dictionary
    Set firstVariables = Cache.Variables

    Set SecondaryDictionary = LLdictionary.Create(ThisWorkbook.Worksheets(SECONDARY_SHEET), 1, 1)
    LinelistStub.SetDictionary SecondaryDictionary

    Cache.Refresh

    Set secondDictionary = Cache.Dictionary
    Set refreshedVariables = Cache.Variables

    Assert.IsFalse firstDictionary Is secondDictionary, "Refresh should pick up the new dictionary reference"
    Assert.IsFalse firstVariables Is refreshedVariables, "Refresh should rebuild the variables helper"
    Exit Sub

Fail:
    FailUnexpectedError Assert, "TestRefreshRehydratesDictionaryChanges"
End Sub

