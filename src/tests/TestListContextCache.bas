Attribute VB_Name = "TestListContextCache"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@ModuleDescription("Unit tests validating the ListContextCache behaviour")
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Assert As Object
Private Cache As IListContextCache

'@section Module lifecycle
'===============================================================================
'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@section Test lifecycle
'===============================================================================
'@TestInitialize
Private Sub TestInitialize()
    Set Cache = ListContextCache.Create
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set Cache = Nothing
End Sub

'@section Tests
'===============================================================================
'@TestMethod("ListContextCache")
Private Sub TestAttachAndRetrieveReferences()
    Dim dictionary As Object
    Dim variables As Object

    Set dictionary = New Collection
    Set variables = New Collection

    Cache.AttachDictionary dictionary
    Cache.AttachVariables variables

    Assert.IsTrue Cache.Dictionary Is dictionary, "Dictionary reference should be preserved"
    Assert.IsTrue Cache.Variables Is variables, "Variables reference should be preserved"
End Sub

'@TestMethod("ListContextCache")
Private Sub TestDictionaryMustBeAttached()
    On Error GoTo ExpectError
    Dim obj As Object
    Set obj = Cache.Dictionary
    Assert.Fail "Retrieving without attachment should raise"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), Err.Number
    Err.Clear
End Sub

'@TestMethod("ListContextCache")
Private Sub TestInvalidateClearsReferences()
    Cache.AttachDictionary New Collection
    Cache.AttachVariables New Collection

    Cache.Invalidate

    On Error GoTo ExpectDictionaryError
    Dim obj As Object
    Set obj = Cache.Dictionary
    Assert.Fail "Dictionary should have been cleared"
    Exit Sub

ExpectDictionaryError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), Err.Number
    Err.Clear

    On Error GoTo ExpectVariablesError
    Set obj = Cache.Variables
    Assert.Fail "Variables should have been cleared"
    Exit Sub

ExpectVariablesError:
    Assert.AreEqual CLng(ProjectError.ObjectNotInitialized), Err.Number
    Err.Clear
End Sub

'@TestMethod("ListContextCache")
Private Sub TestAttachRejectsNothing()
    On Error GoTo ExpectError
    Cache.AttachDictionary Nothing
    Assert.Fail "Attaching Nothing should raise"
    Exit Sub

ExpectError:
    Assert.AreEqual CLng(ProjectError.InvalidArgument), Err.Number
    Err.Clear
End Sub
