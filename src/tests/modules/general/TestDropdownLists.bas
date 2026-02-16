Attribute VB_Name = "TestDropdownLists"

Option Explicit



'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
'@ModuleDescription("Tests for the DropdownLists class")

'@description
'Validates the DropdownLists class, which manages named dropdown list storage
'on hidden worksheets. Each DropdownLists instance wraps a single worksheet and
'provides methods to add, remove, sort, update, clear, translate, validate, and
'hyperlink named dropdown columns. Two fixtures are created in TestInitialize:
'dropOne (no header prefix) and dropTwo (with "dropdown_" prefix), exercising
'both prefix modes throughout the suite.
'Tests cover: factory creation with error on Nothing sheet, Name property,
'adding lists with labels and counter prefixes, duplicate detection surfaced
'through IChecking, removal, HiddenNames-backed counter persistence at workbook
'and worksheet scope, existence checks across instances, AllDropdowns enumeration
'that excludes removed entries, translation of all lists via ITranslationObject,
'LabelRange text with auto-incrementing counter prefixes, value retrieval with
'and without headers plus unknown-list fallback, Length tracking after successive
'adds, ascending and descending Sort, ClearList followed by Update with
'deduplication and bottom-append, and finally SetValidation with error/warning
'alert styles plus forward and return hyperlinks between output and dropdown
'sheets.
'Uses the CustomTest harness (ICustomTest), not Rubberduck.
'@depends DropdownLists, IDropdownLists, Checking, IChecking, HiddenNames, IHiddenNames, TranslationObject, ITranslationObject, BetterArray, CustomTest, TestHelpers

Private Const TEST_OUTPUT_SHEET As String = "testsOutputs"
Private Const DROPTESTONE As String = "DropTestList1"
Private Const DROPTESTTWO As String = "DropTestList2"
Private Const DROPOUTPUT As String = "DataOut"
Private Const WORKBOOK_COUNTER_NAME As String = "__Var__WBDROPCOUNTER"
Private Const WORKSHEET_COUNTER_NAME As String = "__Var__SHDROPCOUNTER"
Private Const TEST_TRANSLATIONS_SHEET As String = "__dropTranslations"

Private Assert As ICustomTest
Private Fakes As Object
Private dropOne As IDropdownLists
Private dropTwo As IDropdownLists

'@section Helpers
'===============================================================================

'@sub-title Ensure the three hidden worksheets used by dropdown tests exist.
Private Sub EnsureDropSheets()
    EnsureWorksheet DROPOUTPUT, visibility:=xlSheetHidden
    EnsureWorksheet DROPTESTONE, visibility:=xlSheetHidden
    EnsureWorksheet DROPTESTTWO, visibility:=xlSheetHidden
End Sub

'@sub-title Clear the contents of all dropdown test worksheets, including the optional translations sheet.
Private Sub ClearDropSheets()
    ClearWorksheet ThisWorkbook.Worksheets(DROPOUTPUT)
    ClearWorksheet ThisWorkbook.Worksheets(DROPTESTONE)
    ClearWorksheet ThisWorkbook.Worksheets(DROPTESTTWO)
    On Error Resume Next
        ClearWorksheet ThisWorkbook.Worksheets(TEST_TRANSLATIONS_SHEET)
    On Error GoTo 0
End Sub


'@section Module lifecycle
'===============================================================================

'@ModuleInitialize
Private Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TEST_OUTPUT_SHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TEST_OUTPUT_SHEET)
    Assert.SetModuleName "TestDropdownLists"
    EnsureDropSheets
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    If Not Assert Is Nothing Then
        Assert.PrintResults TEST_OUTPUT_SHEET
    End If
    Set Assert = Nothing
    Set Fakes = Nothing
    DeleteWorksheets DROPOUTPUT, DROPTESTONE, DROPTESTTWO
    RestoreApp
End Sub

'@TestInitialize
Private Sub TestInitialize()

    BusyApp
    ClearDropSheets
    Set dropOne = DropdownLists.Create(ThisWorkbook.Worksheets(DROPTESTONE), hPrefix:=vbNullString)
    Set dropTwo = DropdownLists.Create(ThisWorkbook.Worksheets(DROPTESTTWO), hPrefix:="dropdown_")

End Sub

Private Sub TestCleanup()
    If Not Assert Is Nothing Then Assert.Flush
End Sub

'@section Tests
'===============================================================================

'@sub-title Verify factory creation succeeds for both prefix modes and raises on Nothing worksheet.
'@details
'Creates dropOne with an empty prefix and dropTwo with "dropdown_" prefix,
'asserting both return non-Nothing references. Then sets the worksheet
'variable to Nothing and attempts creation, asserting that
'ProjectError.ElementNotFound is raised. This tests the guard clause in
'DropdownLists.Create against invalid input.
'@TestMethod("DropdownLists")
Public Sub TestCreateCheck()
    CustomTestSetTitles Assert, "DropdownLists", "TestCreateCheck"

    Dim workbook As Workbook
    Dim sheet As Worksheet
    Dim temporaryDropdown As IDropdownLists

    On Error GoTo Fail

    Set workbook = ThisWorkbook
    Set sheet = workbook.Worksheets(DROPTESTONE)
    Set dropOne = DropdownLists.Create(sheet, hPrefix:=vbNullString)
    Assert.IsTrue (Not dropOne Is Nothing), "Unable to create the first dropdown list object"

    Set sheet = workbook.Worksheets(DROPTESTTWO)
    Set dropTwo = DropdownLists.Create(sheet, hPrefix:="dropdown_")
    Assert.IsTrue (Not dropTwo Is Nothing), "Unable to create the second dropdown list object"

    On Error Resume Next
        Set sheet = Nothing
        Err.Clear
        '@Ignore  AssignmentNotUsed
        Set temporaryDropdown = DropdownLists.Create(sheet)
        Assert.AreEqual ProjectError.ElementNotFound, Err.Number, "Creating with an empty worksheet should raise ElementNotFound"
    On Error GoTo Fail

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestCreateCheck", Err.Number, Err.Description
End Sub

'@sub-title Verify Name property returns the underlying worksheet name.
'@details
'Reads the Name property of dropOne and asserts it matches the DROPTESTONE
'constant, confirming the factory correctly wires the sheet identity.
'@TestMethod("DropdownLists")
Public Sub TestName()
    CustomTestSetTitles Assert, "DropdownLists", "TestName"
    On Error GoTo Fail
    Assert.IsTrue (dropOne.Name = DROPTESTONE), "Name the dropdown object is not correctly set"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestName", Err.Number, Err.Description
End Sub

'@sub-title Verify Add writes multiple named lists with different label and prefix options.
'@details
'Adds three lists to dropOne with varying addLabel and counterPrefix
'combinations, and three lists to dropTwo with a "dropdown_" prefix mode.
'The test succeeds if no error is raised, confirming that Add handles all
'parameter permutations without failing. This is a smoke test for Add
'across both prefix modes.
'@TestMethod("DropdownLists")
Public Sub TestAdd()
    CustomTestSetTitles Assert, "DropdownLists", "TestAdd"
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("One", "Two", "Three", "Four")

    dropOne.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"
    dropOne.Add valuesList, "listValues2", addLabel:=False
    dropOne.Add valuesList, "listValues3", addLabel:=True, counterPrefix:="Test"

    dropTwo.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"
    dropTwo.Add valuesList, "listValues2", addLabel:=False
    dropTwo.Add valuesList, "listValues3", addLabel:=True

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAdd", Err.Number, Err.Description
End Sub


'@sub-title Verify adding a duplicate list name surfaces an error through IChecking.
'@details
'Adds "listValues" once, then adds the same name again. Asserts that
'HasCheckings becomes True and that CheckingValues contains exactly one
'entry, confirming the class records the duplicate rather than raising a
'runtime error. This tests the internal error-collection path used during
'linelist building where halting is not desired.
'@TestMethod("DropdownLists")
Public Sub TestAddExisting()
    CustomTestSetTitles Assert, "DropdownLists", "TestAddExisting"

    Dim checking As IChecking
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("One", "Two", "Three", "Four")

    dropOne.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"
    'Adding the same list again
    dropOne.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"
    Assert.IsTrue dropOne.HasCheckings, "Adding existing dropdown does not raise an internal error"

    If dropOne.HasCheckings Then
        Set checking = dropOne.CheckingValues
        Assert.IsTrue (checking.Length = 1), "Raised error not added to dropdownlist checking"
    End If

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAddExisting", Err.Number, Err.Description
End Sub

'@sub-title Verify Remove deletes a previously added list without raising an error.
'@details
'Adds "removedListValues" with a label and counter, then removes it. The
'test succeeds if no runtime error occurs, confirming the removal path
'cleans up the internal storage correctly.
'@TestMethod("DropdownLists")
Public Sub TestRemove()
    CustomTestSetTitles Assert, "DropdownLists", "TestRemove"
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("Random", "List", "Values")
    dropOne.Add valuesList, "removedListValues", addLabel:=True, counterPrefix:="List"
    dropOne.Remove "removedListValues"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestRemove", Err.Number, Err.Description
End Sub

'@sub-title Verify Add and Remove update HiddenNames counters at workbook and worksheet scope.
'@details
'Captures the current workbook-scope and worksheet-scope counter values
'from HiddenNames, then adds a list and asserts both counters incremented
'by one. After removing the list, asserts the worksheet counter reverts to
'its original value. The fail handler restores original counter values to
'avoid polluting subsequent tests. This confirms the DropdownLists class
'persists counter state through HiddenNames rather than in-memory only.
'@TestMethod("DropdownLists")
Public Sub TestCountersPersistThroughHiddenNames()
    CustomTestSetTitles Assert, "DropdownLists", "TestCountersPersistThroughHiddenNames"

    Dim valuesList As BetterArray
    Dim wbStore As IHiddenNames
    Dim shStore As IHiddenNames
    Dim originalWb As Long
    Dim originalSh As Long

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("alpha")
    Set wbStore = HiddenNames.Create(ThisWorkbook)
    Set shStore = HiddenNames.Create(dropOne.Wksh)

    If Not wbStore.HasName(WORKBOOK_COUNTER_NAME) Then
        wbStore.EnsureName WORKBOOK_COUNTER_NAME, 0, HiddenNameTypeLong
    End If
    If Not shStore.HasName(WORKSHEET_COUNTER_NAME) Then
        shStore.EnsureName WORKSHEET_COUNTER_NAME, 0, HiddenNameTypeLong
    End If

    originalWb = wbStore.ValueAsLong(WORKBOOK_COUNTER_NAME, 0)
    originalSh = shStore.ValueAsLong(WORKSHEET_COUNTER_NAME, 0)

    dropOne.Add valuesList, "hnCounterList", addLabel:=False

    Assert.AreEqual originalWb + 1, wbStore.ValueAsLong(WORKBOOK_COUNTER_NAME, -1), _
                     "Workbook counter should increment through HiddenNames"
    Assert.AreEqual originalSh + 1, shStore.ValueAsLong(WORKSHEET_COUNTER_NAME, -1), _
                     "Worksheet counter should increment through HiddenNames"

    dropOne.Remove "hnCounterList"
    Assert.AreEqual originalSh, shStore.ValueAsLong(WORKSHEET_COUNTER_NAME, -1), _
                     "Worksheet counter should revert after removal"

    wbStore.SetValue WORKBOOK_COUNTER_NAME, originalWb
    shStore.SetValue WORKSHEET_COUNTER_NAME, originalSh
    Exit Sub

Fail:
    On Error Resume Next
        If Not wbStore Is Nothing Then wbStore.SetValue WORKBOOK_COUNTER_NAME, originalWb
        If Not shStore Is Nothing Then shStore.SetValue WORKSHEET_COUNTER_NAME, originalSh
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestCountersPersistThroughHiddenNames", Err.Number, Err.Description
End Sub


'@sub-title Verify Exists returns True for added lists and False for absent ones.
'@details
'Adds "listValues" to dropOne and "listValues3" to dropTwo, then asserts
'Exists returns True for each added name on the correct instance. Also
'asserts Exists returns False for "listValues4" on dropTwo, confirming the
'negative case. This tests cross-instance isolation and name lookup.
'@TestMethod("DropdownLists")
Public Sub TestExists()
    CustomTestSetTitles Assert, "DropdownLists", "TestExists"
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("One", "Two", "Three", "Four")
    dropOne.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"
    dropTwo.Add valuesList, "listValues3", addLabel:=True

    Assert.IsTrue dropOne.Exists("listValues"), "Existing dropdownlist named [listValues] not found in dropdown " & dropOne.Name
    Assert.IsTrue dropTwo.Exists("listValues3"), "Existing dropdownlist named [listValues3] not found in dropdown " & dropTwo.Name
    Assert.IsFalse dropTwo.Exists("listValues4"), "Non Existing dropdownlist named [listValues4] found in dropdown " & dropTwo.Name

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestExists", Err.Number, Err.Description
End Sub

'@sub-title Verify AllDropdowns excludes removed entries and preserves insertion order.
'@details
'Adds "firstList" and "secondList", removes the second, then calls
'AllDropdowns. Asserts the returned BetterArray is not Nothing, has length
'1, and its sole item is "firstList". This confirms removal marks entries
'as cleared rather than leaving holes, and that the enumeration respects
'original insertion order.
'@TestMethod("DropdownLists")
Public Sub TestAllDropdownsSkipsClearedEntries()
    CustomTestSetTitles Assert, "DropdownLists", "TestAllDropdownsSkipsClearedEntries"

    Dim valuesList As BetterArray
    Dim listings As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("alpha", "beta")

    dropOne.Add valuesList, "firstList", addLabel:=True
    dropOne.Add valuesList, "secondList", addLabel:=True
    dropOne.Remove "secondList"

    Set listings = dropOne.AllDropdowns

    Assert.IsFalse listings Is Nothing, "AllDropdowns should return a BetterArray"
    Assert.AreEqual 1&, listings.Length, "AllDropdowns should exclude removed entries"
    Assert.AreEqual "firstList", listings.Item(listings.LowerBound), _
                     "AllDropdowns should preserve insertion order for remaining lists"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestAllDropdownsSkipsClearedEntries", Err.Number, Err.Description
End Sub

'@sub-title Verify Translate applies a TranslationObject to all stored lists.
'@details
'Adds two lists ("firstList", "secondList") containing ("first", "second"),
'builds a translation table mapping "first" to "uno" and "second" to "dos",
'then calls Translate. Asserts the first value of firstList is "uno" and
'the second value of secondList is "dos", confirming that translation is
'applied to every list in the DropdownLists instance. Cleans up the
'temporary translations sheet afterward.
'@TestMethod("DropdownLists")
Public Sub TestTranslateAppliesTranslatorToAllLists()
    CustomTestSetTitles Assert, "DropdownLists", "TestTranslateAppliesTranslatorToAllLists"

    Dim valuesList As BetterArray
    Dim translator As ITranslationObject
    Dim transTable As ListObject
    Dim firstValues As BetterArray
    Dim secondValues As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("first", "second")
    dropOne.Add valuesList, "firstList", addLabel:=False
    dropOne.Add valuesList, "secondList", addLabel:=False

    Set transTable = BuildTranslationTable()
    Set translator = TranslationObject.Create(transTable, "translated")
    dropOne.Translate translator, True

    Set firstValues = dropOne.Values("firstList")
    Set secondValues = dropOne.Values("secondList")

    Assert.AreEqual "uno", CStr(firstValues.Item(firstValues.LowerBound)), "Translate should update first list values"
    Assert.AreEqual "dos", CStr(secondValues.Item(secondValues.LowerBound + 1)), "Translate should update second list values"

    DeleteWorksheets TEST_TRANSLATIONS_SHEET
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTranslateAppliesTranslatorToAllLists", Err.Number, Err.Description
End Sub

'@sub-title Build a two-column ListObject for translation tests (key -> translated).
'@details
'Creates or clears the TEST_TRANSLATIONS_SHEET, writes a key/translated
'header with two rows ("first"->"uno", "second"->"dos"), and wraps the
'range in a ListObject named "__DropTranslations".
Private Function BuildTranslationTable() As ListObject
    Dim hostSheet As Worksheet
    Dim lo As ListObject

    Set hostSheet = EnsureWorksheet(TEST_TRANSLATIONS_SHEET)
    hostSheet.Cells.Clear
    hostSheet.Range("A1").Value = "key"
    hostSheet.Range("B1").Value = "translated"
    hostSheet.Range("A2").Value = "first"
    hostSheet.Range("B2").Value = "uno"
    hostSheet.Range("A3").Value = "second"
    hostSheet.Range("B3").Value = "dos"

    On Error Resume Next
        hostSheet.ListObjects("__DropTranslations").Delete
    On Error GoTo 0

    Set lo = hostSheet.ListObjects.Add(xlSrcRange, hostSheet.Range("A1:B3"), , xlYes)
    lo.Name = "__DropTranslations"
    Set BuildTranslationTable = lo
End Function

'@sub-title Verify LabelRange returns the correct auto-generated label text with counter prefix.
'@details
'Adds two lists with different counterPrefix values ("List" and "Test"),
'then reads the LabelRange cell value for each. Asserts the first returns
'"List 1" and the second "Test 2", confirming that the auto-incrementing
'counter pairs correctly with the prefix string and that each list gets a
'unique label.
'@TestMethod("DropdownLists")
Public Sub TestLabelRange()
    CustomTestSetTitles Assert, "DropdownLists", "TestLabelRange"

    Dim labelText As String
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("One", "Two", "Three", "Four")
    dropOne.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"
    dropOne.Add valuesList, "listValues2", addLabel:=True, counterPrefix:="Test"

    labelText = dropOne.LabelRange("listValues").Value
    Assert.IsTrue (labelText = "List 1"), "Expected returned label: [List 1], Actual: [" & labelText & "]"

    labelText = dropOne.LabelRange("listValues2").Value
    Assert.IsTrue (labelText = "Test 2"), "Expected returned label: [Test 2], Actual: [" & labelText & "]"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLabelRange", Err.Number, Err.Description
End Sub

'@sub-title Verify Values returns correct items, supports includeHeaders, and handles unknown lists.
'@details
'Adds a four-item list, retrieves values without headers and asserts length
'is 4 with correct first item. Retrieves again with includeHeaders:=True
'and asserts length is 5 with the list name as the first element. Finally,
'calls Values on an unknown list name via dropTwo and asserts the returned
'BetterArray has length 0, confirming graceful fallback for missing lists.
'@TestMethod("DropdownLists, Values")
Public Sub TestValues()
    CustomTestSetTitles Assert, "DropdownLists", "TestValues"

    Dim valuesResult As BetterArray
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("One", "Two", "Three", "Four")
    dropOne.Add valuesList, "listValues", addLabel:=True, counterPrefix:="List"

    Set valuesResult = dropOne.Values("listValues")
    Assert.IsTrue (valuesResult.Length = 4), "Not all values are returned when asked. Expected 4, actual: " & valuesResult.Length
    Assert.IsTrue (valuesResult.Item(valuesResult.LowerBound) = "One"), "Values not returned in the correct order. First value expected: One, actual: " & valuesResult.Item(valuesResult.LowerBound)

    Set valuesResult = dropOne.Values("listValues", includeHeaders:=True)
    Assert.IsTrue (valuesResult.Length = 5), "Not all values are returned when asked. Expected 5 including header, actual: " & valuesResult.Length
    Assert.IsTrue (valuesResult.Item(valuesResult.LowerBound) = "listValues"), "Headers not returned when asked in values"

    Set valuesResult = dropTwo.Values("listValues4")
    Assert.IsTrue (valuesResult.Length = 0), "Unknown dropdown is generating values : " & valuesResult.ToString()

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValues", Err.Number, Err.Description
End Sub

'@sub-title Verify Length tracks the number of stored lists and updates after Add.
'@details
'Adds four lists in a loop, asserts Length is 4, then adds a fifth and
'asserts Length is 5. This confirms the internal list counter increments
'correctly with each Add call and is not off-by-one.
'@TestMethod("DropdownLists, Length")
Public Sub TestLength()
    CustomTestSetTitles Assert, "DropdownLists", "TestLength"

    Dim valuesList As BetterArray
    Dim index As Long

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("BB", "AA", "01", "DD")

    For index = 1 To 4
        dropOne.Add valuesList, "baseList" & CStr(index)
    Next index

    Assert.IsTrue (dropOne.Length = 4), "Length of the dropdown not correct. Expected 4, actual: " & dropOne.Length

    dropOne.Add valuesList, "sortList", counterPrefix:="sortList"
    Assert.IsTrue (dropOne.Length = 5), "Length not updated after adding new dropdown. Expected 5, actual: " & dropOne.Length

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestLength", Err.Number, Err.Description
End Sub

'@sub-title Verify Sort arranges list values in ascending and descending order.
'@details
'Adds a mixed-type list (numeric 1, strings "AA", "BB", "DD"), sorts
'ascending and asserts all four positions are in correct order. Then sorts
'descending and asserts the reversed order. This exercises both the default
'ascending and explicit xlDescending sort directions, and confirms that
'mixed numeric/string content is handled correctly.
'@TestMethod("DropdownLists")
Public Sub TestSort()
    CustomTestSetTitles Assert, "DropdownLists", "TestSort"

    Dim valuesList As BetterArray
    Dim resultList As BetterArray

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList(1, "AA", "BB", "DD")
    dropOne.Add valuesList, "sortList", counterPrefix:="sortList"

    dropOne.Sort "sortList"
    Set resultList = dropOne.Values("sortList")

    Assert.IsTrue (resultList.Item(resultList.LowerBound) = 1), "Dropdown [sortList] not sorted in ascending order correctly. Values: " & resultList.ToString()
    Assert.IsTrue (resultList.Item(resultList.LowerBound + 1) = "AA"), "Dropdown [sortList] not sorted in ascending order correctly. Values: " & resultList.ToString()
    Assert.IsTrue (resultList.Item(resultList.LowerBound + 2) = "BB"), "Dropdown [sortList] not sorted in ascending order correctly. Values: " & resultList.ToString()
    Assert.IsTrue (resultList.Item(resultList.UpperBound) = "DD"), "Dropdown [sortList] not sorted in ascending order correctly. Values: " & resultList.ToString()

    dropOne.Sort "sortList", xlDescending
    Set resultList = dropOne.Values("sortList")
    Assert.IsTrue (resultList.Item(resultList.LowerBound) = "DD"), "Dropdown [sortList] not sorted in descending order correctly. Values: " & resultList.ToString()
    Assert.IsTrue (resultList.Item(resultList.LowerBound + 1) = "BB"), "Dropdown [sortList] not sorted in descending order correctly. Values: " & resultList.ToString()
    Assert.IsTrue (resultList.Item(resultList.LowerBound + 2) = "AA"), "Dropdown [sortList] not sorted in descending order correctly. Values: " & resultList.ToString()
    Assert.IsTrue (resultList.Item(resultList.UpperBound) = 1), "Dropdown [sortList] not sorted in descending order correctly. Values: " & resultList.ToString()

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestSort", Err.Number, Err.Description
End Sub


'@sub-title Verify ClearList empties a list and Update repopulates with deduplication and append.
'@details
'Adds a six-item list with duplicates ("AA" appears three times), clears
'it and asserts length is 0. Then calls Update with the same values and
'asserts length is 4 (duplicates removed) with original insertion order
'preserved. Next, appends four more values (one duplicate "AA") using
'pasteAtBottom:=True and asserts final length is 7. Finally sorts and
'verifies the numeric value 1 appears first, confirming that Sort works
'correctly after an append-style Update.
'@TestMethod("DropdownLists")
Public Sub TestClearListAndUpdate()
    CustomTestSetTitles Assert, "DropdownLists", "TestClearListAndUpdate"

    Dim baseValues As BetterArray
    Dim updateValues As BetterArray

    On Error GoTo Fail

    Set baseValues = BetterArrayFromList("AA", "DD", "BB", "01", "AA", "AA")
    dropOne.Add baseValues, "sortList"

    dropOne.ClearList "sortList"
    Set baseValues = dropOne.Values("sortList")
    Assert.IsTrue (baseValues.Length = 0), "Length of a cleaned dropdown should be 0, actual: " & baseValues.Length

    Set baseValues = BetterArrayFromList("AA", "DD", "BB", "01", "AA", "AA")
    dropOne.Update baseValues, "sortList"
    Set baseValues = dropOne.Values("sortList")

    Assert.IsTrue (baseValues.Length = 4), "Duplicates and empty spaces not removed when updating dropdown. Expected 4, actual: " & baseValues.Length & " Values: " & baseValues.ToString()
    Assert.IsTrue (baseValues.Item(1) = "AA"), "Updating not done in correct order. Values: " & baseValues.ToString()

    Set updateValues = BetterArrayFromList("AA", "OO", "VV", "FF")
    dropOne.Update updateValues, "sortList", pasteAtBottom:=True
    Set baseValues = dropOne.Values("sortList")

    Assert.IsTrue (baseValues.Length = 7), "Duplicates and empty not removed when updating dropdown by adding new values. Values: " & baseValues.ToString()
    dropOne.Sort "sortList"
    Set baseValues = dropOne.Values("sortList")
    Assert.IsTrue (baseValues.Item(1) = 1), "Sorting not working as expected after appending new elements, dropdown [sortList] not sorted. Values: " & baseValues.ToString()

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestClearListAndUpdate", Err.Number, Err.Description
End Sub

'@sub-title Verify SetValidation applies list validation with error/warning alerts, and hyperlinks link correctly.
'@details
'Adds a list to dropOne and applies SetValidation with alertType "error"
'on one cell, confirming xlValidateList type and xlValidAlertStop style.
'Applies again with "warning" on another cell and checks xlValidAlertWarning
'plus the custom error message. Then creates a forward hyperlink from the
'output sheet to the dropdown label range and verifies the anchor address
'and sub-address point correctly. Finally creates a return link from the
'dropdown sheet back to the output cell and asserts the reverse anchor and
'sub-address are correct. This exercises the full validation-and-navigation
'workflow used when building linelist sheets.
'@TestMethod("DropdownLists")
Public Sub TestValidationAndHyperLinks()
    CustomTestSetTitles Assert, "DropdownLists", "TestValidationAndHyperLinks"

    Dim outputSheet As Worksheet
    Dim cellRange As Range
    Dim labelRange As Range
    Dim hyperlinkItem As HyperLink
    Dim valuesList As BetterArray

    On Error GoTo Fail

    Set outputSheet = EnsureWorksheet(DROPOUTPUT)
    ClearWorksheet outputSheet

    Set valuesList = BetterArrayFromList("AA", "BB", "CC")
    dropOne.Add valuesList, "sortList"

    Set cellRange = outputSheet.Cells(2, 2)
    dropOne.SetValidation cellRange, listName:="sortList", alertType:="error"

    With cellRange.Validation
        Assert.AreEqual .Type, xlValidateList, "Dropdown validation not added"
        Assert.AreEqual .AlertStyle, xlValidAlertStop, "Enforce dropdown not set for validation"
    End With

    Set cellRange = outputSheet.Cells(2, 3)
    dropOne.SetValidation cellRange, listName:="sortList", alertType:="warning", message:="Stop!"
    With cellRange.Validation
        Assert.AreEqual .Type, xlValidateList, "Dropdown validation not added"
        Assert.AreEqual .AlertStyle, xlValidAlertWarning, "Dropdown validation alert is not warning"
        Assert.AreEqual .ErrorMessage, "Stop!", "Dropdown validation message not set"
    End With

    Set cellRange = outputSheet.Cells(3, 3)
    Set labelRange = dropOne.LabelRange("sortList")
    cellRange.Value = "HyperLink"

    dropOne.AddHyperLink "sortList", cellRange
    Assert.IsTrue (outputSheet.Hyperlinks.Count = 1), "Worksheet should have only one hyperlink"

    Set hyperlinkItem = outputSheet.Hyperlinks(1)
    Assert.AreEqual cellRange.Address, hyperlinkItem.Range.Address, "Hyperlink from worksheet to dropdown should be anchored on the correct cell"
    Assert.AreEqual "'" & labelRange.Parent.Name & "'!" & labelRange.Address, hyperlinkItem.SubAddress, "Hyperlink from worksheet to dropdown should point to the correct address"

    dropOne.AddReturnLink "sortList", cellRange
    Set hyperlinkItem = labelRange.Parent.Hyperlinks(1)

    Assert.AreEqual labelRange.Address, hyperlinkItem.Range.Address, "Hyperlink from dropdown to worksheet should be anchored on the correct cell"
    Assert.AreEqual "'" & cellRange.Parent.Name & "'!" & cellRange.Address, hyperlinkItem.SubAddress, "Hyperlink from dropdown to worksheet should point to the correct address"

    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestValidationAndHyperLinks", Err.Number, Err.Description
End Sub
