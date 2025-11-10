Attribute VB_Name = "TestDropdownLists"

Option Explicit



'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName
'@Folder("CustomTests")
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

Private Sub EnsureDropSheets()
    EnsureWorksheet DROPOUTPUT, visibility:=xlSheetHidden
    EnsureWorksheet DROPTESTONE, visibility:=xlSheetHidden
    EnsureWorksheet DROPTESTTWO, visibility:=xlSheetHidden
End Sub

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

'@TestMethod("DropdownLists")
Public Sub TestName()
    CustomTestSetTitles Assert, "DropdownLists", "TestName"
    On Error GoTo Fail
    Assert.IsTrue (dropOne.Name = DROPTESTONE), "Name the dropdown object is not correctly set"

    Exit Sub
Fail:
    CustomTestLogFailure Assert, "TestName", Err.Number, Err.Description
End Sub

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

'@TestMethod("DropdownLists")
Public Sub TestTranslateAppliesTranslatorToAllLists()
    CustomTestSetTitles Assert, "DropdownLists", "TestTranslateAppliesTranslatorToAllLists"

    Dim valuesList As BetterArray
    Dim translator As ITranslationObject
    Dim transTable As ListObject
    Dim sheet As Worksheet
    Dim listOne As ListObject
    Dim listTwo As ListObject

    On Error GoTo Fail

    Set valuesList = BetterArrayFromList("first", "second")
    dropOne.Add valuesList, "firstList", addLabel:=False
    dropOne.Add valuesList, "secondList", addLabel:=False

    Set transTable = BuildTranslationTable()
    Set translator = TranslationObject.Create(transTable, "translated")
    dropOne.Translate translator, True

    Set sheet = ThisWorkbook.Worksheets(DROPTESTONE)
    Set listOne = FindListObjectFor(sheet, "firstList")
    Set listTwo = FindListObjectFor(sheet, "secondList")

    Assert.IsFalse listOne Is Nothing, "First dropdown listobject not found after translation"
    Assert.IsFalse listTwo Is Nothing, "Second dropdown listobject not found after translation"
    Assert.AreEqual "uno", CStr(listOne.DataBodyRange.Cells(1, 1).Value), "Translate should update first list values"
    Assert.AreEqual "dos", CStr(listTwo.DataBodyRange.Cells(2, 1).Value), "Translate should update second list values"

    DeleteWorksheets TEST_TRANSLATIONS_SHEET
    Exit Sub

Fail:
    CustomTestLogFailure Assert, "TestTranslateAppliesTranslatorToAllLists", Err.Number, Err.Description
End Sub

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

Private Function FindListObjectFor(ByVal hostSheet As Worksheet, ByVal listName As String) As ListObject
    Dim candidate As ListObject
    Dim expectedHeader As String

    expectedHeader = Replace(Trim$(listName), " ", "_")

    For Each candidate In hostSheet.ListObjects
        If Not candidate.HeaderRowRange Is Nothing Then
            If StrComp(CStr(candidate.HeaderRowRange.Cells(1, 1).Value), expectedHeader, vbTextCompare) = 0 Then
                Set FindListObjectFor = candidate
                Exit Function
            End If
        End If
    Next candidate
End Function

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
