Attribute VB_Name = "TestChoiceKeys"
Attribute VB_Description = "Unit tests for ChoiceKeys"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Unit tests for ChoiceKeys")

Option Explicit

'@description
'Drives ChoiceKeys, the one place the export side and the import side build the
'name a custom dropdown is written under on a Choices sheet.
'
'WHAT THESE TESTS ARE GUARDING
'-------------------------------------------------------------------------------
'The two ends used to build the key from two different places and neither was
'the registry name: the export read the ListObject header cell, which holds the
'name with every space turned into an underscore, and the import cut five
'characters off the ListObject name, which leaves the workbook counter attached.
'A dropdown called `contact type` went into the file as
'`__choice_custom_contact_type` and was looked for as
'`__choice_custom_contact type7`.
'
'So the tests that matter are the ones asserting that a name with a space comes
'back with its space, and that the round trip closes. Either would have failed
'against both of the old shapes.
'
'The class holds no state and touches no worksheet, so there is no fixture.
'@depends ChoiceKeys, CustomTest

Private Assert As CustomTest

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "ChoiceKeys"

Private Const CUSTOM_PREFIX As String = "__choice_custom_"


'@section Lifecycle
'===============================================================================

'@sub-title Build the assertion harness.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestChoiceKeys"
End Sub

'@sub-title Print the results.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Put the application into its test state.
'@TestInitialize
Public Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of the test that just ran.
'@TestCleanup
Public Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Factory
'===============================================================================

'@sub-title Create answers a usable instance.
'@TestMethod("ChoiceKeys")
Public Sub FactoryCreatesAnInstance()
    CustomTestSetTitles Assert, TESTMODULE, "FactoryCreatesAnInstance"
    On Error GoTo TestFail

    Dim keys As ChoiceKeys
    Set keys = ChoiceKeys.Create()
    Assert.IsNotNothing keys, "Create should answer an object"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "FactoryCreatesAnInstance", Err.Number, Err.Description
End Sub


'@section Building the key
'===============================================================================

'@sub-title A dropdown name gets the custom prefix and nothing else.
'@TestMethod("ChoiceKeys")
Public Sub TestTheKeyIsThePrefixAndTheName()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheKeyIsThePrefixAndTheName"
    On Error GoTo TestFail

    Assert.AreEqual CUSTOM_PREFIX & "district", _
                    ChoiceKeys.Create().CustomChoiceName("district"), _
                    "The key is the prefix followed by the dropdown name"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheKeyIsThePrefixAndTheName", Err.Number, Err.Description
End Sub

'@sub-title A name carrying a space keeps its space.
'@details
'This is the export half of the fault. The old export read the ListObject header
'cell, which is the name with every space turned into an underscore, so a
'dropdown called `contact type` was written into the file under a name no
'DropdownLists call would ever take back.
'@TestMethod("ChoiceKeys")
Public Sub TestASpaceInTheNameSurvives()
    CustomTestSetTitles Assert, TESTMODULE, "TestASpaceInTheNameSurvives"
    On Error GoTo TestFail

    Dim answer As String

    answer = ChoiceKeys.Create().CustomChoiceName("contact type")

    Assert.AreEqual CUSTOM_PREFIX & "contact type", answer, _
                    "The dropdown name is used as the registry holds it"
    Assert.IsTrue InStr(1, answer, "_type") = 0, _
                  "No space is turned into an underscore"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASpaceInTheNameSurvives", Err.Number, Err.Description
End Sub

'@sub-title An empty dropdown name answers an empty key.
'@details
'LLChoices.Categories reads an empty criteria as "blank cells" rather than as
'"nothing", so handing it a bare prefix would answer rows belonging to no
'dropdown at all.
'@TestMethod("ChoiceKeys")
Public Sub TestAnEmptyNameAnswersAnEmptyKey()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnEmptyNameAnswersAnEmptyKey"
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, ChoiceKeys.Create().CustomChoiceName(vbNullString), _
                    "An empty name never named a dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnEmptyNameAnswersAnEmptyKey", Err.Number, Err.Description
End Sub


'@section Reading a key back
'===============================================================================

'@sub-title A custom key is recognised and anything else is not.
'@TestMethod("ChoiceKeys")
Public Sub TestOnlyACustomKeyIsRecognised()
    CustomTestSetTitles Assert, TESTMODULE, "TestOnlyACustomKeyIsRecognised"
    On Error GoTo TestFail

    Dim keys As ChoiceKeys
    Set keys = ChoiceKeys.Create()

    Assert.IsTrue keys.IsCustomChoice(CUSTOM_PREFIX & "district"), _
                  "A key carrying the prefix is a custom dropdown"
    Assert.IsFalse keys.IsCustomChoice("list_yesno"), _
                   "A choice the setup authored is not a custom dropdown"
    Assert.IsFalse keys.IsCustomChoice(vbNullString), _
                   "An empty key names nothing"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestOnlyACustomKeyIsRecognised", Err.Number, Err.Description
End Sub

'@sub-title The dropdown name comes back out of the key.
'@TestMethod("ChoiceKeys")
Public Sub TestTheNameComesBackOutOfTheKey()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheNameComesBackOutOfTheKey"
    On Error GoTo TestFail

    Assert.AreEqual "district", _
                    ChoiceKeys.Create().ListNameFromChoice(CUSTOM_PREFIX & "district"), _
                    "The inverse gives the dropdown name back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheNameComesBackOutOfTheKey", Err.Number, Err.Description
End Sub

'@sub-title A key with no custom prefix answers an empty name.
'@TestMethod("ChoiceKeys")
Public Sub TestAPlainChoiceAnswersAnEmptyName()
    CustomTestSetTitles Assert, TESTMODULE, "TestAPlainChoiceAnswersAnEmptyName"
    On Error GoTo TestFail

    Assert.AreEqual vbNullString, _
                    ChoiceKeys.Create().ListNameFromChoice("list_yesno"), _
                    "A choice the setup authored came from no dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAPlainChoiceAnswersAnEmptyName", Err.Number, Err.Description
End Sub

'@sub-title The round trip closes, spaces and all.
'@details
'The whole reason the class exists. Whatever the export writes, the import gets
'the same registry name back, so DropdownLists.Update can be handed it.
'@TestMethod("ChoiceKeys")
Public Sub TestTheRoundTripClosesOnANameWithASpace()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheRoundTripClosesOnANameWithASpace"
    On Error GoTo TestFail

    Dim keys As ChoiceKeys
    Dim original As String

    Set keys = ChoiceKeys.Create()
    original = "contact type"

    Assert.AreEqual original, _
                    keys.ListNameFromChoice(keys.CustomChoiceName(original)), _
                    "Name to key to name gives the name back"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheRoundTripClosesOnANameWithASpace", Err.Number, Err.Description
End Sub

'@sub-title The prefix is matched without regard to case.
'@TestMethod("ChoiceKeys")
Public Sub TestThePrefixIsMatchedWithoutCase()
    CustomTestSetTitles Assert, TESTMODULE, "TestThePrefixIsMatchedWithoutCase"
    On Error GoTo TestFail

    Assert.IsTrue ChoiceKeys.Create().IsCustomChoice("__CHOICE_CUSTOM_district"), _
                  "A key written in capitals is still a custom dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestThePrefixIsMatchedWithoutCase", Err.Number, Err.Description
End Sub
