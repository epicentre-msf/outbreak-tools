Attribute VB_Name = "TestEventLinelistSheets"
Attribute VB_Description = "Tests for EventLinelist over built data entry worksheets"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@Folder("CustomTests")
'@ModuleDescription("Tests for EventLinelist over built data entry worksheets")

Option Explicit

'@description
'Drives the four heavy members of EventLinelist over worksheets LLDataEntry
'built: OnSheetChange, OnSelectionChange, UpdateFilterTables and
'UpdateAllListAuto. TestEventLinelist covers the held managers and the flags
'that guard them over hand-seeded sheets. This module covers what those four
'members do on the layout the current builder writes.
'
'THE SHEETS ARE BUILT BY THE PRODUCTION BUILDER, ONCE FOR THE WHOLE MODULE
'-------------------------------------------------------------------------------
'A hand-seeded sheet agrees with whatever the test that wrote it believed. That
'is what let the go-to branch of the HList handler stay dead on every sheet the
'builder makes while the suite reported green (Session 63). So the fixture here
'is one workbook shaped like a designer and one standing in for the linelist,
'and both data entry sheets go through LLDataEntry.Build. The build writes every
'label, dropdown and conditional format of a sheet, so it runs once in
'ModuleInitialize and every test reads and edits what it left. This is the shape
'TestLLDataEntry uses.
'
'THE FIXTURE WORKBOOK CARRIES NO TRANSLATION SHEET, ON PURPOSE
'-------------------------------------------------------------------------------
'The header branch of the HList handler answers a refused edit with
'Warn "MSG_NotModify", and Warn shows a message box whenever the workbook has a
'usable translation sheet. A box in a headless run stops the whole pass on a
'modal. EventLinelist stays quiet when it has no translator, so the linelist
'workbook here is built without a LinelistTranslation worksheet and the header
'test measures the restore alone.
'
'THE FIXTURE FAILS QUIETLY AND EVERY TEST SAYS SO
'-------------------------------------------------------------------------------
'An error escaping ModuleInitialize reaches the VBE as a modal dialog and the
'run comes back with no results file. The setup captures its error into two
'module fields and a guard at the top of every test reports it as that test's
'own failure.
'
'THE SCREEN GOES BACK TO THIS WORKBOOK AFTER ANYTHING ACTIVATES
'-------------------------------------------------------------------------------
'A build freezes the panes of the sheet it wrote, and `LLDataEntry.FreezeHeader`
'activates that sheet to do it, so the linelist workbook holds the screen when
'the build ends. The go-to test activates its sheet too, because the branch
'answers by moving the selection. `CustomTest.PrintResults` writes into a
'worksheet of THIS workbook and raises 1004 while another workbook is in front,
'and a raise inside `ModuleCleanup` is a modal dialog that ends the whole
'headless run with an empty results file. `HandBackTheScreen` runs after the
'build, at the top of `ModuleCleanup`, and on both paths of the go-to test.
'This cost one run on 2026-08-11.
'
'THE BUILT SHEETS ARE UNPROTECTED BEFORE THE TESTS RUN
'-------------------------------------------------------------------------------
'The last statement of a build protects the sheet. Every production caller that
'writes to a data entry sheet unprotects it first through the Passwords manager,
'and the fixture does the same with the manager the build itself used.
'
'THE LIST AUTO COLUMN IS FILLED THE WAY LinelistSpecs FILLS IT
'-------------------------------------------------------------------------------
'LinelistSpecs.AddListAuto marks the variable a list_auto variable draws from
'with "list_auto_origin", and the fixture reaches LinelistSpecs through the
'TestAssign setters, which skip that step. So the fixture marks text_h2 itself
'before the build. VarWriter then writes "text_h2 -- listauto" and the
'sheet-level has_listauto flag, which is what UpdateAllListAuto reads.
'
'THE GEO CASCADE NEEDS A GEOBASE IN BOTH WORKBOOKS
'-------------------------------------------------------------------------------
'`LLdictionary.Prepare` expands a geo variable into geo1 to geo4 columns only
'when it is handed an `LLGeo`, so the designer workbook gets a geobase before
'the dictionary is prepared. `LinelistSpecs` resolves its own geo through the
'worksheet called Geo of that same workbook, which is what lets `VarWriter`
'build the four admin dropdown lists while it writes the geo1 column. The
'linelist workbook gets a geobase of its own, because `EventLinelist.EnsureGeo`
'reads the Geo worksheet of the workbook it was configured with and a
'generation copies the designer's across. Both come from `GeoTestFixture`
'under `helpers`.
'@depends EventLinelist, LLDataEntry, Linelist, LinelistSpecs, LLdictionary,
'  LLVariables, LLSheets, LLFormat, LLChoices, LLGeo, FormulaData, DropdownLists,
'  HiddenNames, Passwords, CustomTest, DictionaryTestFixture, ChoicesTestFixture,
'  LLFormatTestFixture, FormulaTestFixture, PasswordsTestFixture, GeoTestFixture

Private Assert As CustomTest
Private SpecsWkb As Workbook
Private OutWkb As Workbook
Private Dict As LLdictionary
Private Vars As LLVariables
Private Geo As LLGeo
Private Specs As LinelistSpecs
Private LL As Linelist
Private SheetInfo As LLSheets
Private Guard As Passwords
Private Sut As EventLinelist
Private SetupError As Long
Private SetupMessage As String

Private Const TESTOUTPUTSHEET As String = "testsOutputs"
Private Const TESTMODULE As String = "EventLinelist"
Private Const DICTIONARY_SHEET As String = "DictFixture"
Private Const CHOICES_SHEET As String = "ChoicesFixture"
Private Const PASSWORD_SHEET As String = "__pass"

'The two data entry sheets carry the names the dictionary gives them, as the
'worksheets a generation creates do.
Private Const VLIST_SHEET As String = "vlist1D-sheet1"
Private Const HLIST_SHEET As String = "hlist2D-sheet1"

'The worksheets the cached managers of EventLinelist read. The dictionary copy
'is what lets EnsureDictionary and EnsureVariables answer, which the editable
'label branch needs.
Private Const SHEET_DROPDOWN_LISTS As String = "__dropdown_lists"
Private Const SHEET_CUSTOM_CHOICE As String = "LLSHEET_CustomChoice"
Private Const SHEET_CUSTOM_PIVOT As String = "LLSHEET_CustomPivotTable"
Private Const SHEET_DICTIONARY As String = "Dictionary"
Private Const GEO_SHEET As String = "Geo"

'The four dropdown lists the geo cascade rewrites. VarWriter creates them when
'it writes the geo1 column, and every geo sheet of a linelist shares them.
Private Const LIST_ADMIN2 As String = "admin2"
Private Const LIST_ADMIN3 As String = "admin3"
Private Const LIST_ADMIN4 As String = "admin4"

'The hidden name the list auto branch raises, on the SHEET that was edited.
'One workbook-level flag could not say which sheet needed rebuilding.
Private Const LISTAUTO_FLAG As String = "update_listauto"

'The variables of the dictionary fixture each branch is driven through.
'  text_h2       carries "yes" in the Editable Label column and stands in as the
'                origin of the list_auto variable lauto_man_h2
'  int_h2        carries no editable label, so its label is left alone
'  choi_mult_h2  carries the choice_multiple control
'  hid_h2        carries the hidden status, so the build hides its column
'  ed_var_v1     the editable label of the VList sheet
'  choi_mult_v1  the choice_multiple control of the VList sheet
Private Const EDITABLE_VAR As String = "text_h2"
Private Const PLAIN_VAR As String = "int_h2"
Private Const CHOICE_VAR As String = "choi_mult_h2"
Private Const HIDDEN_VAR As String = "hid_h2"
Private Const VLIST_EDITABLE_VAR As String = "ed_var_v1"
Private Const VLIST_CHOICE_VAR As String = "choi_mult_v1"

'AppendGeoLines turns the geo variable geo_h2 into twelve rows. The four the
'cascade drives are adm1_ to adm4_, carrying the controls geo1 to geo4, and
'they sit in that order on the built sheet.
Private Const GEO_VAR_ADM1 As String = "adm1_geo_h2"
Private Const GEO_VAR_ADM2 As String = "adm2_geo_h2"
Private Const GEO_VAR_ADM3 As String = "adm3_geo_h2"
Private Const GEO_VAR_ADM4 As String = "adm4_geo_h2"

'What the geobase fixture holds: three admin1 values, two admin2 under each,
'two admin3 under each of those.
Private Const GEO_ADM1_VALUE As String = "P1"
Private Const GEO_ADM2_OF_P1 As String = "D1"
Private Const GEO_ADM2_SECOND As String = "D2"
Private Const GEO_ADM3_OF_D1 As String = "C1"
Private Const GEO_ADM4_OF_C1 As String = "V1"

'The main label int_h2 arrives with. The test that writes over its label row
'reads this back out of the dictionary.
Private Const PLAIN_VAR_LABEL As String = "Integer on hlist2D"

'One data row of the HList table per test that writes to it, counted from the
'header row. The three list auto rows sit at the top and stay next to each
'other: the reader of a list auto column walks down from the first data cell
'and stops at the first empty one.
Private Const ROW_LISTAUTO_ONE As Long = 1
Private Const ROW_LISTAUTO_TWO As Long = 2
Private Const ROW_LISTAUTO_THREE As Long = 3
Private Const ROW_CHOICE_APPEND As Long = 4
Private Const ROW_CHOICE_TWICE As Long = 5
Private Const ROW_CHOICE_UNHELD As Long = 6
Private Const ROW_FILTER_VISIBLE As Long = 7
Private Const ROW_FILTER_HIDDEN As Long = 8
Private Const ROW_FILTER_HIDDENCOL As Long = 9
Private Const ROW_GEO_CHANGE As Long = 10
Private Const ROW_GEO_SELECTION As Long = 11

'The dropdown an automatic list fills is named after the ORIGIN variable, the
'one users type into, plus this suffix. VarWriter builds the name from the
'control details of the list_auto variable (`lauto_man_h2` names `text_h2`
'there) and UpdateListAutoSheet builds the same name from the column it is
'walking, so both sides land on `text_h2_lauto`. The suffix is the constant
'both of them carry.
Private Const LISTAUTO_SUFFIX As String = "_lauto"

'What the tests type into the sheets. Each one is unique, so a test reads its
'own writing back out of a table every other test has also written to.
Private Const LISTAUTO_VALUE_ONE As String = "listauto-alpha"
Private Const LISTAUTO_VALUE_TWO As String = "listauto-beta"
Private Const FILTER_VISIBLE_VALUE As String = "filter-visible-row"
Private Const FILTER_HIDDEN_VALUE As String = "filter-hidden-row"
Private Const FILTER_HIDDENCOL_VALUE As String = "filter-hidden-column"
Private Const SCRIBBLED_HEADER As String = "typed over the header"
Private Const NEW_LABEL As String = "A label the user typed"

'The separator the multiple choice toggle falls back to when the control string
'carries none.
Private Const CHOICE_SEPARATOR As String = ", "


'@section Lifecycle
'===============================================================================

'@sub-title Build both workbooks, both data entry sheets and the service, once.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleInitialize
Public Sub ModuleInitialize()
    BusyApp
    EnsureWorksheet TESTOUTPUTSHEET, clearSheet:=False
    Set Assert = CustomTest.Create(ThisWorkbook, TESTOUTPUTSHEET)
    Assert.SetModuleName "TestEventLinelistSheets"

    SetupError = 0
    SetupMessage = vbNullString

    'The four steps are read one at a time. Read once at the end, the
    'answer was whatever Err held by then, and every `On Error` statement
    'inside a later call clears Err: a build that raised on its first step
    'reported SetupError = 0, FixtureReady called the fixture ready, and
    'every test then ran against a workbook carrying no data entry table.
    On Error Resume Next
        BuildFixture
        NoteSetupStep "BuildFixture"
        BuildBothSheets
        NoteSetupStep "BuildBothSheets"
        OpenTheBuiltSheets
        NoteSetupStep "OpenTheBuiltSheets"
        Set Sut = EventLinelist.Create(OutWkb)
        NoteSetupStep "EventLinelist.Create"
    On Error GoTo 0

    'A build freezes the panes of the sheet it wrote, and freezing them
    'activates that sheet, so the linelist workbook is the active one when the
    'build ends. The harness writes its results into a worksheet of THIS
    'workbook, and it raises 1004 when another workbook holds the screen.
    HandBackTheScreen
End Sub

'@sub-title Print results and drop both workbooks.
'@details
'This routine is Public because the harness calls it by name through
'Application.Run.
'@ModuleCleanup
Public Sub ModuleCleanup()
    'PrintResults writes into a worksheet of this workbook and raises 1004 when
    'another workbook holds the screen. A raise in a lifecycle hook is a modal,
    'and a modal is the whole run.
    HandBackTheScreen

    If Not Assert Is Nothing Then
        Assert.PrintResults TESTOUTPUTSHEET
    End If

    On Error Resume Next
        If Not SpecsWkb Is Nothing Then DeleteWorkbook SpecsWkb
        If Not OutWkb Is Nothing Then DeleteWorkbook OutWkb
    On Error GoTo 0

    Set Sut = Nothing
    Set Guard = Nothing
    Set Geo = Nothing
    Set SheetInfo = Nothing
    Set LL = Nothing
    Set Specs = Nothing
    Set Vars = Nothing
    Set Dict = Nothing
    Set SpecsWkb = Nothing
    Set OutWkb = Nothing

    RestoreApp
    Set Assert = Nothing
End Sub

'@sub-title Suppress screen updates before each test.
'@details
'There is no BeginTest call here on purpose. BeginTest opens the checking with
'whatever titles are pending at that moment, and the Flush in TestCleanup has
'just reset those to the default, so every result of the module would be filed
'under the default label.
'@TestInitialize
Private Sub TestInitialize()
    BusyApp
End Sub

'@sub-title Flush the results of the test that just ran.
'@TestCleanup
Private Sub TestCleanup()
    If Not Assert Is Nothing Then
        Assert.Flush
    End If
End Sub


'@section Fixture
'===============================================================================

'@sub-title Build the designer-shaped workbook and the linelist workbook.
'@details
'The linelist workbook takes the two data entry sheets, the three worksheets
'the cached managers read, a copy of the dictionary and the list auto flag a
'generated linelist carries. It takes no translation worksheet, which is what
'keeps every message box of the event surface shut.
Private Sub BuildFixture()
    Dim transStub As TranslationObject
    Dim design As LLFormat
    Dim formatSheet As Worksheet
    Dim formData As FormulaData
    Dim formulaSheet As Worksheet
    Dim choicesObj As LLChoices
    Dim designerGeoSheet As Worksheet

    Set SpecsWkb = NewWorkbook()
    Set OutWkb = NewWorkbook()

    'The geobase of the designer. Prepare expands a geo variable into geo1 to
    'geo4 columns only when it is handed an LLGeo, so this is built before the
    'dictionary is prepared. LinelistSpecs resolves its own geo through the
    'worksheet called Geo of the designer workbook, which is this one, and that
    'is what lets VarWriter build the four admin dropdown lists.
    Set designerGeoSheet = GeoTestFixture.PrepareGeoFixture(GEO_SHEET, SpecsWkb, _
                                                            withData:=True)
    Set Geo = LLGeo.Create(designerGeoSheet)

    DictionaryTestFixture.PrepareDictionaryFixture DICTIONARY_SHEET, SpecsWkb
    Set Dict = LLdictionary.Create(SpecsWkb.Worksheets(DICTIONARY_SHEET), 1, 1)
    Dict.Prepare geoObject:=Geo

    'What LinelistSpecs.AddListAuto writes in a generation. The TestAssign
    'setters below hand the dictionary over without running it.
    Dict.AddColumn "list auto"
    Set Vars = LLVariables.Create(Dict)
    Vars.SetValue varName:=EDITABLE_VAR, colName:="list auto", _
                  newValue:="list_auto_origin", onEmpty:=True

    ChoicesTestFixture.PrepareChoicesFixture CHOICES_SHEET, SpecsWkb
    Set choicesObj = LLChoices.Create(SpecsWkb.Worksheets(CHOICES_SHEET), 1, 1)

    'The tags AddLabel and the GoTo dropdown read. A tag the table does not
    'carry answers with the tag itself, which is what makes the three worksheet
    'names below resolve.
    Set transStub = BuildTranslationObject(SpecsWkb, "ENG", _
        Array(Array("MSG_Calculated", "Calculated"), _
              Array("MSG_Mandatory", "Mandatory"), _
              Array("MSG_CustomChoice", "Custom choice"), _
              Array("MSG_GoToSec", "Go to section")))

    Set formatSheet = LLFormatTestFixture.PrepareLLFormatFixture("LLFormatFixture", SpecsWkb)
    Set design = LLFormat.Create(formatSheet)
    Set formulaSheet = FormulaTestFixture.PrepareFormulaFixtureSheet("FormulaFixture", outwb:=SpecsWkb)
    Set formData = FormulaData.Create(formulaSheet)

    EnsureSpecsSheets SpecsWkb
    PasswordsTestFixture.PreparePasswordsFixture PASSWORD_SHEET, SpecsWkb

    EnsureWorksheet VLIST_SHEET, OutWkb, clearSheet:=True
    EnsureWorksheet HLIST_SHEET, OutWkb, clearSheet:=True
    EnsureWorksheet SHEET_DROPDOWN_LISTS, OutWkb, clearSheet:=True, visibility:=xlSheetHidden
    EnsureWorksheet SHEET_CUSTOM_CHOICE, OutWkb, clearSheet:=True, visibility:=xlSheetHidden
    EnsureWorksheet SHEET_CUSTOM_PIVOT, OutWkb, clearSheet:=True, visibility:=xlSheetHidden

    'The dictionary copy a generated linelist carries. EnsureDictionary reads
    'the worksheet called Dictionary of the host workbook, and the editable
    'label branch writes the new label back into it.
    DictionaryTestFixture.PrepareDictionaryFixture SHEET_DICTIONARY, OutWkb

    'The geobase copy a generated linelist carries. EnsureGeo reads the
    'worksheet called Geo of the host workbook, and the cascade asks it for the
    'children of the value the user picked. A generation copies the designer's
    'geo sheet across, so the two carry the same geography.
    GeoTestFixture.PrepareGeoFixture GEO_SHEET, OutWkb, withData:=True

    Set Specs = LinelistSpecs.Create(SpecsWkb)
    Specs.TestAssignDictionary Dict
    Specs.TestAssignDesignFormat design
    Specs.TestAssignTransObject transStub
    Specs.TestAssignFormulaData formData
    Specs.TestAssignChoices choicesObj
    Specs.TestAssignLLWorkbook OutWkb

    Set LL = Linelist.Create(Specs)
    Set SheetInfo = LLSheets.Create(Dict)
    Set Guard = Specs.Password
End Sub

'@sub-title Build the VList sheet and the HList sheet, once each.
Private Sub BuildBothSheets()
    Dim builder As LLDataEntry

    Set builder = LLDataEntry.Create(LLDataEntryLayerVList, VLIST_SHEET, LL, SheetInfo)
    builder.Build

    Set builder = LLDataEntry.Create(LLDataEntryLayerHList, HLIST_SHEET, LL, SheetInfo)
    builder.Build
End Sub

'@sub-title Give the screen back to the workbook the harness writes into.
'@details
'Two things here activate the linelist workbook: a build freezes the panes of
'the sheet it wrote, and the go-to test brings its sheet to the front so the
'branch can move the selection. `CustomTest.PrintResults` writes into a
'worksheet of this workbook and raises 1004 while another workbook holds the
'screen, and a raise inside a lifecycle hook is a modal dialog that stops the
'whole headless run. Every path that activates anything ends here.
Private Sub HandBackTheScreen()
    On Error Resume Next
        ThisWorkbook.Activate
    On Error GoTo 0
End Sub

'@sub-title Take the protection off both built sheets.
'@details
'The last statement of a build protects the sheet, and every test below types
'into one. The manager here is the manager the build used, so the sheets are
'opened the way a button opens them.
Private Sub OpenTheBuiltSheets()
    Guard.UnProtect OutWkb.Worksheets(VLIST_SHEET)
    Guard.UnProtect OutWkb.Worksheets(HLIST_SHEET)
End Sub

'@sub-title Keep the first setup failure, naming the step that raised it.
'@details
'Called after every step of the build while On Error Resume Next is on.
'The first failure is the one that matters: the steps after it run on a
'workbook the failed step never finished.
'@param stepName String. The step just run.
Private Sub NoteSetupStep(ByVal stepName As String)
    If SetupError <> 0 Then Exit Sub
    If Err.Number = 0 Then Exit Sub

    SetupError = Err.Number
    SetupMessage = stepName & " - " & Err.Description
    Err.Clear
End Sub

'@fun-title Report a fixture that could not be built, once per test.
'@param testName String. The test asking.
'@return Boolean. True when the fixture is there.
Private Function FixtureReady(ByVal testName As String) As Boolean
    If (SetupError = 0) And (Not Sut Is Nothing) Then
        FixtureReady = True
        Exit Function
    End If

    CustomTestLogFailure Assert, testName, SetupError, _
                         "The fixture could not be built - " & SetupMessage
End Function


'@section Fixture readers
'===============================================================================

'@fun-title The HList data entry sheet.
'@return Worksheet. The built HList sheet.
Private Function HListSheet() As Worksheet
    Set HListSheet = OutWkb.Worksheets(HLIST_SHEET)
End Function

'@fun-title The VList data entry sheet.
'@return Worksheet. The built VList sheet.
Private Function VListSheet() As Worksheet
    Set VListSheet = OutWkb.Worksheets(VLIST_SHEET)
End Function

'@fun-title The data entry table of the HList sheet.
'@details
'The first ListObject is what EventLinelist itself reads, so this reads the
'same one.
'@return ListObject. The data entry table.
Private Function HListLo() As ListObject
    Set HListLo = HListSheet().ListObjects(1)
End Function

'@fun-title The hidden names of one built sheet.
'@param sh Worksheet. The sheet whose store is wanted.
'@return HiddenNames. A store over that sheet.
Private Function StoreOf(ByVal sh As Worksheet) As HiddenNames
    Set StoreOf = HiddenNames.Create(sh)
End Function

'@fun-title The worksheet column one variable heads on the HList sheet.
'@param varName String. The variable name in the header row.
'@return Long. The worksheet column, or 0 when the header is absent.
Private Function HeaderColumnOf(ByVal varName As String) As Long
    Dim hRng As Range
    Dim idx As Long

    Set hRng = HListLo().HeaderRowRange

    For idx = 1 To hRng.Columns.Count
        If CStr(hRng.Cells(1, idx).Value) = varName Then
            HeaderColumnOf = hRng.Cells(1, idx).Column
            Exit Function
        End If
    Next idx
End Function

'@fun-title One data cell of the HList table.
'@param varName String. The variable heading the column.
'@param rowOffset Long. Rows below the header row.
'@return Range. The cell.
Private Function HListCell(ByVal varName As String, ByVal rowOffset As Long) As Range
    Set HListCell = HListSheet().Cells(HListLo().HeaderRowRange.Row + rowOffset, _
                                       HeaderColumnOf(varName))
End Function

'@fun-title The header cell of one variable of the HList table.
'@param varName String. The variable heading the column.
'@return Range. The header cell.
Private Function HListHeaderCell(ByVal varName As String) As Range
    Set HListHeaderCell = HListCell(varName, 0)
End Function

'@fun-title The label cell of one variable of the HList table.
'@details
'The builder writes the main label one row above the header row and names that
'cell after the variable.
'@param varName String. The variable heading the column.
'@return Range. The label cell.
Private Function HListLabelCell(ByVal varName As String) As Range
    Set HListLabelCell = HListCell(varName, -1)
End Function

'@fun-title The value cell of one variable of the VList sheet.
'@details
'The builder names the value cell after the variable, and writes the main label
'in the cell to its left.
'@param varName String. The variable.
'@return Range. The value cell.
Private Function VListCell(ByVal varName As String) As Range
    Set VListCell = VListSheet().Range(varName)
End Function

'@fun-title The filtered companion of the HList sheet.
'@return Worksheet. The sheet the data entry sheet names as its filtered copy.
Private Function FilteredSheet() As Worksheet
    Set FilteredSheet = OutWkb.Worksheets(StoreOf(HListSheet()).ValueAsString("filtered_sheet"))
End Function

'@fun-title Whether the filtered table carries one value anywhere.
'@param wanted String. The value looked for.
'@return Boolean. True when a cell of the filtered table holds it.
Private Function FilteredHolds(ByVal wanted As String) As Boolean
    Dim filtLo As ListObject
    Dim bodyRng As Range
    Dim cellRng As Range

    Set filtLo = FilteredSheet().ListObjects(1)
    Set bodyRng = filtLo.DataBodyRange
    If bodyRng Is Nothing Then Exit Function

    For Each cellRng In bodyRng.Cells
        If CStr(cellRng.Value) = wanted Then
            FilteredHolds = True
            Exit Function
        End If
    Next cellRng
End Function

'@fun-title The values one dropdown list holds.
'@param listName String. The name of the list.
'@return BetterArray. The values of the list.
Private Function DropdownValues(ByVal listName As String) As BetterArray
    Dim drop As DropdownLists

    Set drop = DropdownLists.Create(OutWkb.Worksheets(SHEET_DROPDOWN_LISTS))
    Set DropdownValues = drop.Values(listName)
End Function

'@fun-title How many times one value shows up in a list.
'@param list BetterArray. The list to walk.
'@param wanted String. The value counted.
'@return Long. The number of entries equal to that value.
Private Function CountOf(ByVal list As BetterArray, ByVal wanted As String) As Long
    Dim idx As Long

    If list Is Nothing Then Exit Function
    If list.Length = 0 Then Exit Function

    For idx = list.LowerBound To list.UpperBound
        If CStr(list.Item(idx)) = wanted Then CountOf = CountOf + 1
    Next idx
End Function

'@sub-title Write the three list auto values into the origin column.
'@details
'The reader of a list auto column walks down from the first data cell and stops
'at the first empty one, so the three rows are next to each other and start at
'the top of the table. Two of them hold the same value, which is what the
'unique test reads.
Private Sub SeedListAutoColumn()
    HListCell(EDITABLE_VAR, ROW_LISTAUTO_ONE).Value = LISTAUTO_VALUE_ONE
    HListCell(EDITABLE_VAR, ROW_LISTAUTO_TWO).Value = LISTAUTO_VALUE_ONE
    HListCell(EDITABLE_VAR, ROW_LISTAUTO_THREE).Value = LISTAUTO_VALUE_TWO
End Sub

'@fun-title The first section name written above the table of the HList sheet.
'@details
'The go-to branch searches the row three above the header row, which is where
'the section builder writes the main section names.
'@return String. A section name the sheet carries, or an empty string.
Private Function FirstSectionName() As String
    Dim sectionRng As Range
    Dim idx As Long
    Dim cellValue As String

    Set sectionRng = HListLo().HeaderRowRange.Offset(-3)

    For idx = 1 To sectionRng.Columns.Count
        cellValue = CStr(sectionRng.Cells(1, idx).Value)
        If LenB(cellValue) > 0 Then
            FirstSectionName = cellValue
            Exit Function
        End If
    Next idx
End Function


'@section OnSheetChange -- the HList sheet
'===============================================================================

'@sub-title An edit of the header row is refused and the name put back.
'@details
'The restore reads the name off the label cell above the header, which is the
'cell the builder names after the variable. Session 63 moved this branch below
'the go-to test; before that a raise in the go-to branch stopped it running at
'all.
'@TestMethod("EventLinelist")
Public Sub TestAHeaderEditIsPutBack()
    CustomTestSetTitles Assert, TESTMODULE, "TestAHeaderEditIsPutBack"
    If Not FixtureReady("TestAHeaderEditIsPutBack") Then Exit Sub
    On Error GoTo TestFail

    Dim headerRng As Range
    Dim restored As String

    'A column is resolved by reading the header row, so the header goes back
    'whatever the handler did with it. A header left holding the scribble takes
    'that column away from every test that runs after this one.
    Set headerRng = HListHeaderCell(PLAIN_VAR)
    headerRng.Value = SCRIBBLED_HEADER

    Sut.OnSheetChange HListSheet(), headerRng

    restored = CStr(headerRng.Value)
    headerRng.Value = PLAIN_VAR

    Assert.AreEqual PLAIN_VAR, restored, _
                    "An edit of the header row should be put back to the variable name"

    Exit Sub
TestFail:
    On Error Resume Next
        If Not headerRng Is Nothing Then headerRng.Value = PLAIN_VAR
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAHeaderEditIsPutBack", Err.Number, Err.Description
End Sub

'@sub-title An editable label reaches the dictionary of the linelist.
'@details
'The handler strips the sub label and the line break the builder joined it
'with, then writes what is left back as the main label. The reader here is the
'dictionary the class holds, which is the copy carried by the linelist
'workbook.
'@TestMethod("EventLinelist")
Public Sub TestAnEditableLabelReachesTheDictionary()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnEditableLabelReachesTheDictionary"
    If Not FixtureReady("TestAnEditableLabelReachesTheDictionary") Then Exit Sub
    On Error GoTo TestFail

    Dim labelRng As Range
    Dim subLabel As String

    subLabel = Sut.Variables.Value(varName:=EDITABLE_VAR, colName:="sub label")

    Set labelRng = HListLabelCell(EDITABLE_VAR)
    labelRng.Value = NEW_LABEL & Chr(10) & subLabel

    Sut.OnSheetChange HListSheet(), labelRng

    Assert.AreEqual NEW_LABEL, _
                    Sut.Variables.Value(varName:=EDITABLE_VAR, colName:="main label"), _
                    "An editable label should reach the main label of the dictionary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnEditableLabelReachesTheDictionary", Err.Number, Err.Description
End Sub

'@sub-title A label the dictionary does not mark editable is left alone.
'@TestMethod("EventLinelist")
Public Sub TestALabelThatIsNotEditableIsLeftAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestALabelThatIsNotEditableIsLeftAlone"
    If Not FixtureReady("TestALabelThatIsNotEditableIsLeftAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim labelRng As Range
    Set labelRng = HListLabelCell(PLAIN_VAR)
    labelRng.Value = NEW_LABEL

    Sut.OnSheetChange HListSheet(), labelRng

    Assert.AreEqual PLAIN_VAR_LABEL, _
                    Sut.Variables.Value(varName:=PLAIN_VAR, colName:="main label"), _
                    "A variable with no editable label should keep the label it was given"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestALabelThatIsNotEditableIsLeftAlone", Err.Number, Err.Description
End Sub

'@sub-title An edit of a list auto column raises the flag of that sheet.
'@details
'The flag is what the deactivate handler reads to decide whether the dropdowns
'of the sheet being left want rebuilding. It lives on the sheet, so an edit
'landing on one sheet cannot send another sheet off to rebuild.
'@TestMethod("EventLinelist")
Public Sub TestAListAutoEditRaisesTheWorkbookFlag()
    CustomTestSetTitles Assert, TESTMODULE, "TestAListAutoEditRaisesTheWorkbookFlag"
    If Not FixtureReady("TestAListAutoEditRaisesTheWorkbookFlag") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range
    Dim sheetNames As HiddenNames

    Set sheetNames = HiddenNames.Create(HListSheet())
    sheetNames.EnsureName LISTAUTO_FLAG, "no", HiddenNameTypeString
    sheetNames.SetValue LISTAUTO_FLAG, "no"

    Set valueRng = HListCell(EDITABLE_VAR, ROW_LISTAUTO_ONE)
    valueRng.Value = LISTAUTO_VALUE_ONE

    Sut.OnSheetChange HListSheet(), valueRng

    Assert.AreEqual "yes", sheetNames.ValueAsString(LISTAUTO_FLAG), _
                    "An edit of a list auto column should raise the flag of that sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAListAutoEditRaisesTheWorkbookFlag", Err.Number, Err.Description
End Sub

'@sub-title A second choice is added to the one the cell already held.
'@details
'The value the cell held comes from what OnSelectionChange stored when the user
'landed on it. The two calls here are the two events a user raises by picking a
'second entry from a multiple choice dropdown.
'@TestMethod("EventLinelist")
Public Sub TestAMultipleChoiceEditAppendsToWhatTheCellHeld()
    CustomTestSetTitles Assert, TESTMODULE, "TestAMultipleChoiceEditAppendsToWhatTheCellHeld"
    If Not FixtureReady("TestAMultipleChoiceEditAppendsToWhatTheCellHeld") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range
    Set valueRng = HListCell(CHOICE_VAR, ROW_CHOICE_APPEND)

    valueRng.Value = "A"
    Sut.OnSelectionChange HListSheet(), valueRng

    valueRng.Value = "B"
    Sut.OnSheetChange HListSheet(), valueRng

    Assert.AreEqual "A" & CHOICE_SEPARATOR & "B", CStr(valueRng.Value), _
                    "A second choice should be added to the one the cell held"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAMultipleChoiceEditAppendsToWhatTheCellHeld", Err.Number, Err.Description
End Sub

'@sub-title Picking the same choice twice leaves the cell as it stood.
'@TestMethod("EventLinelist")
Public Sub TestTheSameChoiceTwiceLeavesTheCellAsItStood()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheSameChoiceTwiceLeavesTheCellAsItStood"
    If Not FixtureReady("TestTheSameChoiceTwiceLeavesTheCellAsItStood") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range
    Set valueRng = HListCell(CHOICE_VAR, ROW_CHOICE_TWICE)

    valueRng.Value = "A"
    Sut.OnSelectionChange HListSheet(), valueRng

    valueRng.Value = "A"
    Sut.OnSheetChange HListSheet(), valueRng

    Assert.AreEqual "A", CStr(valueRng.Value), _
                    "The same choice picked twice should leave one entry in the cell"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheSameChoiceTwiceLeavesTheCellAsItStood", Err.Number, Err.Description
End Sub

'@sub-title The go-to dropdown moves the selection to the section picked.
'@details
'The builder writes the go-to dropdown in the first cell of the sheet and
'stores its caption in the sheet store, and every entry it offers carries that
'caption as a prefix. The branch strips the prefix and looks the section name
'up in the row three above the header. Activating the sheet is what lets the
'branch move the selection, so the test reads the active cell back.
'@TestMethod("EventLinelist")
Public Sub TestTheGoToDropdownMovesToTheSection()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheGoToDropdownMovesToTheSection"
    If Not FixtureReady("TestTheGoToDropdownMovesToTheSection") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim goToRng As Range
    Dim store As HiddenNames
    Dim caption As String
    Dim sectionName As String
    Dim heldCaption As Variant
    Dim landedOn As String

    Set sh = HListSheet()
    Set store = StoreOf(sh)
    caption = store.ValueAsString(store.ValueAsString("table_name") & "_go_to_section")
    sectionName = FirstSectionName()

    Set goToRng = sh.Cells(1, 1)
    heldCaption = goToRng.Value

    'The branch answers by moving the selection, so the sheet has to be in
    'front for the move to land. The screen goes back below, on both paths.
    OutWkb.Activate
    sh.Activate

    goToRng.Value = caption & ": " & sectionName
    Sut.OnSheetChange sh, goToRng

    landedOn = CStr(ActiveCell.Value)

    goToRng.Value = heldCaption
    HandBackTheScreen

    Assert.IsTrue LenB(caption) > 0, _
                  "A built HList sheet should store the caption of its go-to dropdown"
    Assert.AreEqual sectionName, landedOn, _
                    "Picking a section in the go-to dropdown should move to that section"

    Exit Sub
TestFail:
    HandBackTheScreen
    CustomTestLogFailure Assert, "TestTheGoToDropdownMovesToTheSection", Err.Number, Err.Description
End Sub


'@section OnSheetChange -- the geo cascade
'===============================================================================
'The four admin columns sit next to each other on the built sheet, because
'AppendGeoLines gives adm2 to adm4 the column index of adm1 plus one, two and
'three. Every branch below reads the columns to its left and writes the
'dropdown list of the level under it.

'@sub-title Picking an admin1 value fills the dropdown of the level under it.
'@TestMethod("EventLinelist")
Public Sub TestAGeoEditFillsTheDropdownOfTheLevelBelow()
    CustomTestSetTitles Assert, TESTMODULE, "TestAGeoEditFillsTheDropdownOfTheLevelBelow"
    If Not FixtureReady("TestAGeoEditFillsTheDropdownOfTheLevelBelow") Then Exit Sub
    On Error GoTo TestFail

    Dim adm1Rng As Range
    Dim listValues As BetterArray

    Set adm1Rng = HListCell(GEO_VAR_ADM1, ROW_GEO_CHANGE)
    adm1Rng.Value = GEO_ADM1_VALUE

    Sut.OnSheetChange HListSheet(), adm1Rng

    Set listValues = DropdownValues(LIST_ADMIN2)

    Assert.IsTrue CountOf(listValues, GEO_ADM2_OF_P1) > 0, _
                  "The admin2 dropdown should hold the first child of the value picked"
    Assert.IsTrue CountOf(listValues, GEO_ADM2_SECOND) > 0, _
                  "The admin2 dropdown should hold every child of the value picked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAGeoEditFillsTheDropdownOfTheLevelBelow", Err.Number, Err.Description
End Sub

'@sub-title Picking an admin1 value empties the three levels under it.
'@details
'The three cells to the right of the edited one are the lower admin levels, and
'what they held belongs to the admin1 value that was there before.
'@TestMethod("EventLinelist")
Public Sub TestAGeoEditClearsTheLevelsUnderIt()
    CustomTestSetTitles Assert, TESTMODULE, "TestAGeoEditClearsTheLevelsUnderIt"
    If Not FixtureReady("TestAGeoEditClearsTheLevelsUnderIt") Then Exit Sub
    On Error GoTo TestFail

    Dim adm1Rng As Range

    Set adm1Rng = HListCell(GEO_VAR_ADM1, ROW_GEO_CHANGE)

    HListCell(GEO_VAR_ADM2, ROW_GEO_CHANGE).Value = GEO_ADM2_OF_P1
    HListCell(GEO_VAR_ADM3, ROW_GEO_CHANGE).Value = GEO_ADM3_OF_D1
    HListCell(GEO_VAR_ADM4, ROW_GEO_CHANGE).Value = GEO_ADM4_OF_C1

    adm1Rng.Value = GEO_ADM1_VALUE
    Sut.OnSheetChange HListSheet(), adm1Rng

    Assert.AreEqual vbNullString, _
                    CStr(HListCell(GEO_VAR_ADM2, ROW_GEO_CHANGE).Value), _
                    "An admin1 edit should empty the admin2 cell beside it"
    Assert.AreEqual vbNullString, _
                    CStr(HListCell(GEO_VAR_ADM3, ROW_GEO_CHANGE).Value), _
                    "An admin1 edit should empty the admin3 cell beside it"
    Assert.AreEqual vbNullString, _
                    CStr(HListCell(GEO_VAR_ADM4, ROW_GEO_CHANGE).Value), _
                    "An admin1 edit should empty the admin4 cell beside it"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAGeoEditClearsTheLevelsUnderIt", Err.Number, Err.Description
End Sub

'@sub-title Picking an admin2 value fills the admin3 dropdown.
'@details
'This level reads two cells: the admin1 value to its left and the value just
'picked, because an admin2 name is unique under its parent alone.
'@TestMethod("EventLinelist")
Public Sub TestASecondLevelGeoEditFillsTheThirdDropdown()
    CustomTestSetTitles Assert, TESTMODULE, "TestASecondLevelGeoEditFillsTheThirdDropdown"
    If Not FixtureReady("TestASecondLevelGeoEditFillsTheThirdDropdown") Then Exit Sub
    On Error GoTo TestFail

    Dim adm2Rng As Range

    HListCell(GEO_VAR_ADM1, ROW_GEO_CHANGE).Value = GEO_ADM1_VALUE
    Set adm2Rng = HListCell(GEO_VAR_ADM2, ROW_GEO_CHANGE)
    adm2Rng.Value = GEO_ADM2_OF_P1

    Sut.OnSheetChange HListSheet(), adm2Rng

    Assert.IsTrue CountOf(DropdownValues(LIST_ADMIN3), GEO_ADM3_OF_D1) > 0, _
                  "The admin3 dropdown should hold the children of the admin2 value picked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASecondLevelGeoEditFillsTheThirdDropdown", Err.Number, Err.Description
End Sub

'@sub-title Picking an admin3 value fills the admin4 dropdown.
'@TestMethod("EventLinelist")
Public Sub TestAThirdLevelGeoEditFillsTheFourthDropdown()
    CustomTestSetTitles Assert, TESTMODULE, "TestAThirdLevelGeoEditFillsTheFourthDropdown"
    If Not FixtureReady("TestAThirdLevelGeoEditFillsTheFourthDropdown") Then Exit Sub
    On Error GoTo TestFail

    Dim adm3Rng As Range

    HListCell(GEO_VAR_ADM1, ROW_GEO_CHANGE).Value = GEO_ADM1_VALUE
    HListCell(GEO_VAR_ADM2, ROW_GEO_CHANGE).Value = GEO_ADM2_OF_P1
    Set adm3Rng = HListCell(GEO_VAR_ADM3, ROW_GEO_CHANGE)
    adm3Rng.Value = GEO_ADM3_OF_D1

    Sut.OnSheetChange HListSheet(), adm3Rng

    Assert.IsTrue CountOf(DropdownValues(LIST_ADMIN4), GEO_ADM4_OF_C1) > 0, _
                  "The admin4 dropdown should hold the children of the admin3 value picked"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAThirdLevelGeoEditFillsTheFourthDropdown", Err.Number, Err.Description
End Sub

'@sub-title An empty geo cell leaves the dropdowns as they stand.
'@details
'Deleting the value of a geo cell reaches the handler the same way typing one
'does, and there is no parent to read children from.
'@TestMethod("EventLinelist")
Public Sub TestAnEmptiedGeoCellLeavesTheDropdownAlone()
    CustomTestSetTitles Assert, TESTMODULE, "TestAnEmptiedGeoCellLeavesTheDropdownAlone"
    If Not FixtureReady("TestAnEmptiedGeoCellLeavesTheDropdownAlone") Then Exit Sub
    On Error GoTo TestFail

    Dim adm1Rng As Range
    Dim heldCount As Long

    Set adm1Rng = HListCell(GEO_VAR_ADM1, ROW_GEO_CHANGE)
    adm1Rng.Value = GEO_ADM1_VALUE
    Sut.OnSheetChange HListSheet(), adm1Rng
    heldCount = DropdownValues(LIST_ADMIN2).Length

    adm1Rng.ClearContents
    Sut.OnSheetChange HListSheet(), adm1Rng

    Assert.AreEqual heldCount, DropdownValues(LIST_ADMIN2).Length, _
                    "Emptying a geo cell should leave the dropdown under it as it stood"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAnEmptiedGeoCellLeavesTheDropdownAlone", Err.Number, Err.Description
End Sub


'@section OnSelectionChange -- the geo cascade
'===============================================================================

'@sub-title Landing on an admin2 cell rebuilds its dropdown from its parent.
'@details
'This is what makes the dropdown of a row show the right children when the user
'walks back to a row filled earlier: the change handler wrote the list for
'whichever row was edited last, and arriving on a cell rewrites it for this one.
'@TestMethod("EventLinelist")
Public Sub TestLandingOnAGeoCellRebuildsItsDropdown()
    CustomTestSetTitles Assert, TESTMODULE, "TestLandingOnAGeoCellRebuildsItsDropdown"
    If Not FixtureReady("TestLandingOnAGeoCellRebuildsItsDropdown") Then Exit Sub
    On Error GoTo TestFail

    Dim adm2Rng As Range

    'The row this test walks back to carries P1, and the dropdown is left
    'holding the children of another admin1 value first.
    HListCell(GEO_VAR_ADM1, ROW_GEO_SELECTION).Value = GEO_ADM1_VALUE
    Set adm2Rng = HListCell(GEO_VAR_ADM2, ROW_GEO_SELECTION)

    Sut.OnSelectionChange HListSheet(), adm2Rng

    Assert.IsTrue CountOf(DropdownValues(LIST_ADMIN2), GEO_ADM2_OF_P1) > 0, _
                  "Landing on an admin2 cell should fill its dropdown from the row's admin1 value"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLandingOnAGeoCellRebuildsItsDropdown", Err.Number, Err.Description
End Sub

'@sub-title Landing on an admin3 cell reads both levels above it.
'@TestMethod("EventLinelist")
Public Sub TestLandingOnAThirdLevelCellReadsBothParents()
    CustomTestSetTitles Assert, TESTMODULE, "TestLandingOnAThirdLevelCellReadsBothParents"
    If Not FixtureReady("TestLandingOnAThirdLevelCellReadsBothParents") Then Exit Sub
    On Error GoTo TestFail

    Dim adm3Rng As Range

    HListCell(GEO_VAR_ADM1, ROW_GEO_SELECTION).Value = GEO_ADM1_VALUE
    HListCell(GEO_VAR_ADM2, ROW_GEO_SELECTION).Value = GEO_ADM2_OF_P1
    Set adm3Rng = HListCell(GEO_VAR_ADM3, ROW_GEO_SELECTION)

    Sut.OnSelectionChange HListSheet(), adm3Rng

    Assert.IsTrue CountOf(DropdownValues(LIST_ADMIN3), GEO_ADM3_OF_D1) > 0, _
                  "Landing on an admin3 cell should fill its dropdown from the two levels above"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestLandingOnAThirdLevelCellReadsBothParents", Err.Number, Err.Description
End Sub

'@sub-title The build gives the geo variable its four admin columns.
'@details
'The cascade rests on this: the dictionary carried one variable with a geo
'control, and Prepare turned it into four columns carrying geo1 to geo4, in
'that order and next to each other.
'@TestMethod("EventLinelist")
Public Sub TestTheBuiltSheetCarriesTheFourAdminColumns()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBuiltSheetCarriesTheFourAdminColumns"
    If Not FixtureReady("TestTheBuiltSheetCarriesTheFourAdminColumns") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Dim adm1Column As Long

    Set store = StoreOf(HListSheet())
    adm1Column = HeaderColumnOf(GEO_VAR_ADM1)

    Assert.AreEqual "geo1", store.ValueAsString(GEO_VAR_ADM1 & " -- control"), _
                    "The first admin column should carry the geo1 control"
    Assert.AreEqual "geo4", store.ValueAsString(GEO_VAR_ADM4 & " -- control"), _
                    "The fourth admin column should carry the geo4 control"
    Assert.AreEqual adm1Column + 1, HeaderColumnOf(GEO_VAR_ADM2), _
                    "The admin2 column should sit beside the admin1 column"
    Assert.AreEqual adm1Column + 3, HeaderColumnOf(GEO_VAR_ADM4), _
                    "The four admin columns should sit next to each other"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBuiltSheetCarriesTheFourAdminColumns", Err.Number, Err.Description
End Sub


'@section OnSheetChange -- the VList sheet
'===============================================================================

'@sub-title An editable label of a VList sheet reaches the dictionary.
'@details
'The VList branch reads the variable name off the cell to the right of the
'label, which is the value cell the builder names after the variable.
'@TestMethod("EventLinelist")
Public Sub TestAVListEditableLabelReachesTheDictionary()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListEditableLabelReachesTheDictionary"
    If Not FixtureReady("TestAVListEditableLabelReachesTheDictionary") Then Exit Sub
    On Error GoTo TestFail

    Dim labelRng As Range
    Dim subLabel As String

    subLabel = Sut.Variables.Value(varName:=VLIST_EDITABLE_VAR, colName:="sub label")

    Set labelRng = VListCell(VLIST_EDITABLE_VAR).Offset(, -1)
    labelRng.Value = NEW_LABEL & Chr(10) & subLabel

    Sut.OnSheetChange VListSheet(), labelRng

    Assert.AreEqual NEW_LABEL, _
                    Sut.Variables.Value(varName:=VLIST_EDITABLE_VAR, colName:="main label"), _
                    "An editable label of a VList sheet should reach the dictionary"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListEditableLabelReachesTheDictionary", Err.Number, Err.Description
End Sub

'@sub-title A choice typed into a VList value cell is kept as typed.
'@details
'This states what the current builder gives, and it is worth reading before it
'is taken as the wanted behaviour. The VList branch reads the control string
'from the cell to the RIGHT of the value cell, and the builder writes nothing
'there: it writes the label to the left of the value cell and stores every
'control string as a worksheet hidden name. So the multiple choice toggle
'never runs on a VList sheet, and a second pick replaces the first instead of
'being added to it. The HList branch of the same edit reads the control string
'out of the sheet store and does add it, which
'TestAMultipleChoiceEditAppendsToWhatTheCellHeld measures.
'
'Where the VList control string should live is a design call, so this test
'holds the behaviour as it stands. A fix turns this test red, which is what
'should happen.
'@TestMethod("EventLinelist")
Public Sub TestAVListChoiceEditIsKeptAsTyped()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVListChoiceEditIsKeptAsTyped"
    If Not FixtureReady("TestAVListChoiceEditIsKeptAsTyped") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range
    Set valueRng = VListCell(VLIST_CHOICE_VAR)

    valueRng.Value = "A"
    Sut.OnSelectionChange VListSheet(), valueRng

    valueRng.Value = "B"
    Sut.OnSheetChange VListSheet(), valueRng

    Assert.AreEqual "B", CStr(valueRng.Value), _
                    "A VList value cell has no control string beside it, so the " & _
                    "second choice replaces the first"
    Assert.AreEqual vbNullString, CStr(valueRng.Offset(, 1).Value), _
                    "The cell the VList branch reads the control string from is empty"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVListChoiceEditIsKeptAsTyped", Err.Number, Err.Description
End Sub


'@section OnSelectionChange
'===============================================================================

'@sub-title A selection of several cells drops the held value.
'@details
'One held value can stand for one cell only. After a multi-cell selection the
'toggle has no reading to add to, so the entry the user typed is what the cell
'keeps.
'@TestMethod("EventLinelist")
Public Sub TestASelectionOfSeveralCellsDropsTheHeldValue()
    CustomTestSetTitles Assert, TESTMODULE, "TestASelectionOfSeveralCellsDropsTheHeldValue"
    If Not FixtureReady("TestASelectionOfSeveralCellsDropsTheHeldValue") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim valueRng As Range

    Set sh = HListSheet()
    Set valueRng = HListCell(CHOICE_VAR, ROW_CHOICE_UNHELD)

    valueRng.Value = "A"
    Sut.OnSelectionChange sh, sh.Range(valueRng, valueRng.Offset(1))

    valueRng.Value = "B"
    Sut.OnSheetChange sh, valueRng

    Assert.AreEqual "B", CStr(valueRng.Value), _
                    "With no held reading the toggle should keep the entry typed"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASelectionOfSeveralCellsDropsTheHeldValue", Err.Number, Err.Description
End Sub

'@sub-title A selection above the table is answered and changes nothing.
'@details
'Worksheet_SelectionChange is written into the code module of every HList
'sheet, so the handler is reached by an arrow key anywhere on the sheet,
'including the label rows above the table.
'@TestMethod("EventLinelist")
Public Sub TestASelectionAboveTheTableChangesNothing()
    CustomTestSetTitles Assert, TESTMODULE, "TestASelectionAboveTheTableChangesNothing"
    If Not FixtureReady("TestASelectionAboveTheTableChangesNothing") Then Exit Sub
    On Error GoTo TestFail

    Dim sh As Worksheet
    Dim labelRng As Range
    Dim heldLabel As Variant

    Set sh = HListSheet()
    Set labelRng = HListLabelCell(CHOICE_VAR)
    heldLabel = labelRng.Value

    Sut.OnSelectionChange sh, labelRng

    Assert.AreEqual CStr(heldLabel), CStr(labelRng.Value), _
                    "A selection above the table should leave the label as it stood"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestASelectionAboveTheTableChangesNothing", Err.Number, Err.Description
End Sub


'@section UpdateFilterTables
'===============================================================================

'@sub-title A visible row holding a value reaches the filtered table.
'@details
'The assertions read values rather than a row count. How many rows the table
'carries depends on what the computed columns of the fixture answer on an
'untouched row, and a case_when variable answers its default there, so a count
'states something about the dictionary fixture instead of about the rewrite.
'@TestMethod("EventLinelist")
Public Sub TestAVisibleFilledRowReachesTheFilteredTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAVisibleFilledRowReachesTheFilteredTable"
    If Not FixtureReady("TestAVisibleFilledRowReachesTheFilteredTable") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range

    Set valueRng = HListCell(PLAIN_VAR, ROW_FILTER_VISIBLE)
    valueRng.Value = FILTER_VISIBLE_VALUE

    Sut.UpdateFilterTables calculate:=False

    Assert.IsTrue FilteredHolds(FILTER_VISIBLE_VALUE), _
                  "A visible row holding a value should reach the filtered table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAVisibleFilledRowReachesTheFilteredTable", Err.Number, Err.Description
End Sub

'@sub-title A value taken out of the source leaves the filtered table too.
'@details
'Each pass rewrites the whole filtered table from the source, so a value the
'user deleted is gone from the copy after the next pass.
'@TestMethod("EventLinelist")
Public Sub TestAClearedValueLeavesTheFilteredTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAClearedValueLeavesTheFilteredTable"
    If Not FixtureReady("TestAClearedValueLeavesTheFilteredTable") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range

    Set valueRng = HListCell(PLAIN_VAR, ROW_FILTER_VISIBLE)
    valueRng.Value = FILTER_VISIBLE_VALUE
    Sut.UpdateFilterTables calculate:=False

    valueRng.ClearContents
    Sut.UpdateFilterTables calculate:=False

    Assert.IsFalse FilteredHolds(FILTER_VISIBLE_VALUE), _
                   "A value cleared in the source should be gone from the filtered table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAClearedValueLeavesTheFilteredTable", Err.Number, Err.Description
End Sub

'@sub-title A hidden row is left out of the filtered table.
'@details
'The row is shown again before the assertions run, so the rest of the module
'reads a sheet with nothing hidden.
'@TestMethod("EventLinelist")
Public Sub TestAHiddenRowIsLeftOutOfTheFilteredTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAHiddenRowIsLeftOutOfTheFilteredTable"
    If Not FixtureReady("TestAHiddenRowIsLeftOutOfTheFilteredTable") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range
    Dim reached As Boolean

    Set valueRng = HListCell(PLAIN_VAR, ROW_FILTER_HIDDEN)
    valueRng.Value = FILTER_HIDDEN_VALUE

    valueRng.EntireRow.Hidden = True
    Sut.UpdateFilterTables calculate:=False
    reached = FilteredHolds(FILTER_HIDDEN_VALUE)
    valueRng.EntireRow.Hidden = False

    Assert.IsFalse reached, _
                   "A hidden row should be left out of the filtered table"

    Exit Sub
TestFail:
    On Error Resume Next
        HListCell(PLAIN_VAR, ROW_FILTER_HIDDEN).EntireRow.Hidden = False
    On Error GoTo 0
    CustomTestLogFailure Assert, "TestAHiddenRowIsLeftOutOfTheFilteredTable", Err.Number, Err.Description
End Sub

'@sub-title A hidden column keeps its values in the filtered table.
'@details
'The visibility of a row is read over its whole row on purpose. Show/hide hides
'variable columns, and the filtered table has to carry the values of those
'columns. The column read here is hidden by the build itself, because the
'dictionary gives that variable the hidden status.
'@TestMethod("EventLinelist")
Public Sub TestAHiddenColumnKeepsItsValuesInTheFilteredTable()
    CustomTestSetTitles Assert, TESTMODULE, "TestAHiddenColumnKeepsItsValuesInTheFilteredTable"
    If Not FixtureReady("TestAHiddenColumnKeepsItsValuesInTheFilteredTable") Then Exit Sub
    On Error GoTo TestFail

    Dim valueRng As Range

    Set valueRng = HListCell(HIDDEN_VAR, ROW_FILTER_HIDDENCOL)
    valueRng.Value = FILTER_HIDDENCOL_VALUE

    Sut.UpdateFilterTables calculate:=False

    Assert.IsTrue valueRng.EntireColumn.Hidden, _
                  "The build should hide the column of a variable the dictionary hides"
    Assert.IsTrue FilteredHolds(FILTER_HIDDENCOL_VALUE), _
                  "A hidden column should keep its values in the filtered table"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAHiddenColumnKeepsItsValuesInTheFilteredTable", Err.Number, Err.Description
End Sub


'@section UpdateAllListAuto
'===============================================================================

'@sub-title The dropdown of a list auto column takes the values typed in it.
'@details
'The dropdown of a list auto variable is named after the variable it draws
'from, which is the name VarWriter reads out of the control details column.
'@TestMethod("EventLinelist")
Public Sub TestAListAutoDropdownTakesTheColumnValues()
    CustomTestSetTitles Assert, TESTMODULE, "TestAListAutoDropdownTakesTheColumnValues"
    If Not FixtureReady("TestAListAutoDropdownTakesTheColumnValues") Then Exit Sub
    On Error GoTo TestFail

    Dim listValues As BetterArray

    SeedListAutoColumn
    Sut.UpdateAllListAuto

    Set listValues = DropdownValues(EDITABLE_VAR & LISTAUTO_SUFFIX)

    Assert.IsTrue CountOf(listValues, LISTAUTO_VALUE_ONE) > 0, _
                  "A value typed into a list auto column should reach its dropdown"
    Assert.IsTrue CountOf(listValues, LISTAUTO_VALUE_TWO) > 0, _
                  "Every value of a list auto column should reach its dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAListAutoDropdownTakesTheColumnValues", Err.Number, Err.Description
End Sub

'@sub-title A value typed twice shows up once in the list auto dropdown.
'@TestMethod("EventLinelist")
Public Sub TestAListAutoDropdownKeepsOneEntryPerValue()
    CustomTestSetTitles Assert, TESTMODULE, "TestAListAutoDropdownKeepsOneEntryPerValue"
    If Not FixtureReady("TestAListAutoDropdownKeepsOneEntryPerValue") Then Exit Sub
    On Error GoTo TestFail

    SeedListAutoColumn
    Sut.UpdateAllListAuto

    Assert.AreEqual 1, CountOf(DropdownValues(EDITABLE_VAR & LISTAUTO_SUFFIX), _
                              LISTAUTO_VALUE_ONE), _
                    "A value typed into two rows should show up once in the dropdown"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestAListAutoDropdownKeepsOneEntryPerValue", Err.Number, Err.Description
End Sub

'@sub-title The sheet-level list auto flag is what the walk reads first.
'@details
'A sheet holding no list auto variable is left after one read of its store.
'The built HList sheet carries the flag because the fixture marks the origin
'variable the way LinelistSpecs does.
'@TestMethod("EventLinelist")
Public Sub TestTheBuiltSheetCarriesTheListAutoFlag()
    CustomTestSetTitles Assert, TESTMODULE, "TestTheBuiltSheetCarriesTheListAutoFlag"
    If Not FixtureReady("TestTheBuiltSheetCarriesTheListAutoFlag") Then Exit Sub
    On Error GoTo TestFail

    Dim store As HiddenNames
    Set store = StoreOf(HListSheet())

    Assert.AreEqual "yes", store.ValueAsString("has_listauto"), _
                    "A sheet holding a list auto origin should carry the sheet-level flag"
    Assert.AreEqual "list_auto_origin", _
                    store.ValueAsString(EDITABLE_VAR & " -- listauto"), _
                    "The origin variable should carry its list auto mark on the sheet"

    Exit Sub
TestFail:
    CustomTestLogFailure Assert, "TestTheBuiltSheetCarriesTheListAutoFlag", Err.Number, Err.Description
End Sub
