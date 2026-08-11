VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DropdownLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A wrapper for dropdowns"

'@PredeclaredId
'@Folder("General")
'@ModuleDescription("A wrapper for dropdowns")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, MissingAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@class DropdownLists
'@description
'Manages the dropdown lists of one worksheet. Each list is a one-column
'ListObject and carries a workbook-level name. Add creates a list, Remove
'deletes it, ClearList empties it, Update rewrites its values, and Sort and
'Translate reorder and relabel the entries. SetValidation puts a list behind a
'range of cells, and AddHyperLink and AddReturnLink link the dropdown worksheet
'with the cells that use it.
'
'THE REGISTRY
'-------------------------------------------------------------------------------
'Column C of the worksheet holds the dropdown name and column D the name of its
'ListObject, both under a header on row 2. Every public member reads those two
'columns once as a block and answers from that copy for the rest of the call.
'The copy is dropped when the call ends, so each call reads the columns again.
'@depends HiddenNames, BetterArray, Checking
'@version 1.0 (2026-02-09)


Private Type TDropdownLists
    sh As Worksheet
    hprefix As String
    hasChecks As Boolean
    initialised As Boolean
    wbStored As HiddenNames
    shStored As HiddenNames
    sealed As Boolean
    registryLoaded As Boolean
    registryRows As Long
    registryBlock As Variant
End Type

Private this As TDropdownLists
Private Const CLASSNAME As String = "DropdownLists"
Private Const DEFAULTNAMEPREFIX As String = "choice_"
Private Const WORKBOOKCOUNTERNAME As String = "__Var__WBDROPCOUNTER"  'Workbook-level dropdown counter
Private Const WORKSHEETCOUNTERNAME As String = "__Var__SHDROPCOUNTER" 'Worksheet-level dropdown counter
Private Const DROPDOWNINITIALSEDNAME As String = "__Var__DROPINIT"
Private Const LISTNAMECOLUMNHEADER As String =  "__listName__"
Private Const LONAMECOLUMNHEADER As String = "__loName__"
Private Const WORKSHEETLANGUAGENAME As String = "__Var__LANGNAME"

'increment by default the counter on the number of dropdowns by 1. Everytime a
'dropdown is added to the current worksheet. The counter is saved somewhere the
'first cell fo the worksheet. It starts at 1.
'dropdowns in the current dropdown worksheet.
Private Const DROPDOWNCOUNTERINCREMENT As Byte = 1

'Where all the dropdowns should start, sometimes there could be a need
'for header for the dropdown (like in custom dropdowns, so we shift one line)
Private Const DROPDOWNSTARTROW As Integer = 2

'The registry that maps a dropdown name to its ListObject name. The dropdown
'name is in column C, the ListObject name in column D. Both columns carry a
'header on DROPDOWNSTARTROW.
Private Const REGISTRYNAMECOLUMN As Long = 3
Private Const REGISTRYLONAMECOLUMN As Long = 4

'Field positions inside one row of the registry block held in memory
Private Const REGISTRYNAMEITEM As Long = 1
Private Const REGISTRYLONAMEITEM As Long = 2

'Row 1 of the registry block is the column header, so data starts on row 2
Private Const REGISTRYFIRSTDATAITEM As Long = 2

'Counter start values: the workbook counter starts at 1, the worksheet one at 0
Private Const WORKBOOKCOUNTERSTART As Long = 1
Private Const WORKSHEETCOUNTERSTART As Long = 0


'checking objects
Private checkCounter As Long
Private internalChecks As Checking

'@section Factory
'===============================================================================

'@label:create
'@sub-title Create a new DropdownLists instance
'@details
'Initialises a DropdownLists object on the given worksheet. Each
'dropdown is stored as a one-column ListObject; the header prefix
'distinguishes multiple dropdown worksheets in the same workbook.
'On first use the worksheet is formatted (gridlines removed, registry
'columns created) and an initialisation flag is stored via HiddenNames.
'@param outsh Worksheet. The worksheet to host the dropdown ListObjects.
'@param hprefix Optional String. Prefix for ListObject header names. Defaults to "dropdown_".
'@return DropdownLists. The new instance.
'@throws ProjectError.ElementNotFound When outsh is Nothing.
'@export
Public Function Create(ByVal outsh As Worksheet, Optional ByVal hprefix As String = "dropdown_") As DropdownLists

    If outsh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Dropdown worksheet not found"
    End If

    Dim newDropdown As DropdownLists
    Set newDropdown = New DropdownLists

    Set newDropdown.Wksh = outsh
    newDropdown.headerPrefix = hprefix
    newDropdown.Seal
    Set Create = newDropdown
End Function

'@section Elements
'===============================================================================

'@label:wksh
'@prop-title Worksheet hosting the dropdown lists
'@return Worksheet. The host worksheet.
Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

'@label:wksh-set
'@prop-title Assign the worksheet and initialise the dropdown registry
'@details
'On first assignment, formats the worksheet (removes gridlines, creates
'internal registry columns) and stores an initialisation flag. Resets
'cached HiddenNames stores on every assignment.
'@param outsh Worksheet. The worksheet to host dropdown ListObjects.
Public Property Set Wksh(ByVal outsh As Worksheet)
    GuardNotSealed "Wksh"

    Dim store As HiddenNames
    Set store = HiddenNames.Create(outsh)
    this.initialised = (store.ValueAsString(DROPDOWNINITIALSEDNAME) = "Yes")

    If (Not this.initialised) Then

       'Format the worksheet before creating dropdowns
        RemoveGridLines outsh

        outsh.Range("A:D").Font.Color = vbWhite
        outsh.Range("A:D").Interior.Color = vbWhite

        outsh.Cells(DROPDOWNSTARTROW, REGISTRYNAMECOLUMN).Value = LISTNAMECOLUMNHEADER
        outsh.Cells(DROPDOWNSTARTROW, REGISTRYLONAMECOLUMN).Value = LONAMECOLUMNHEADER

        this.initialised = True
        store.EnsureName DROPDOWNINITIALSEDNAME, "Yes", HiddenNameTypeString
    End If

    Set this.sh = outsh
    Set this.shStored = Nothing
    Set this.wbStored = Nothing
    this.registryLoaded = False
End Property


'@label:all-dropdowns
'@prop-title All registered dropdown names
'@details
'Scans the internal registry column and collects every dropdown name
'into a BetterArray.
'@return BetterArray. The dropdown name list.
Public Property Get AllDropdowns() As BetterArray
    Dim temp As BetterArray
    Dim listName As String
    Dim rowIndex As Long

    Set temp = New BetterArray
    temp.LowerBound = 1

    RefreshRegistry

    For rowIndex = REGISTRYFIRSTDATAITEM To this.registryRows
        listName = RegistryItem(rowIndex, REGISTRYNAMEITEM)
        If LenB(Trim$(listName)) > 0 Then temp.Push listName
    Next rowIndex

    Set AllDropdowns = temp.Clone()
End Property


'@label:header-prefix
'@prop-title Header prefix used for ListObject names
'@return String. The current prefix.
Public Property Get headerPrefix() As String
    headerPrefix = this.hprefix
End Property

'@label:header-prefix-set
'@prop-title Assign the header prefix
'@param hprefix String. The new prefix value.
Public Property Let headerPrefix(ByVal hprefix As String)
    GuardNotSealed "headerPrefix"
    this.hprefix = hprefix
End Property

'@label:name
'@prop-title Name of the dropdown object
'@details
'Returns the name of the worksheet that hosts the dropdown lists.
'@return String. The worksheet name.
Public Property Get Name() As String
    Name = Wksh.Name
End Property


'@section Registry
'===============================================================================

'@label:refresh-registry
'@sub-title Read the dropdown registry again from the worksheet
'@details
'A routine that changes the registry calls this once at its start. One call
'reads the two registry columns a single time, and every lookup after that
'runs in memory. The copy is dropped first because another DropdownLists
'object can hold the same worksheet and change the registry in between.
'
'SetValidation and AddLink read the registry and never write it, so they
'call EnsureRegistry and keep the copy. One SetupPreparation.Prepare makes
'56 SetValidation calls, and each forced reload costs two worksheet
'crossings for bookkeeping alone.
Private Sub RefreshRegistry()
    this.registryLoaded = False
    EnsureRegistry
End Sub

'@label:ensure-registry
'@sub-title Load the registry when no copy is held
Private Sub EnsureRegistry()
    If this.registryLoaded Then Exit Sub
    LoadRegistry
End Sub

'@label:load-registry
'@sub-title Read the two registry columns into memory in one block
'@details
'Reads from the header row down to the last used row of the dropdown
'name column. Row 1 of the block is the header, so dropdown entries
'start at REGISTRYFIRSTDATAITEM.
Private Sub LoadRegistry()
    Dim sh As Worksheet
    Dim lastRow As Long

    this.registryLoaded = True
    this.registryRows = 0
    this.registryBlock = Empty

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    lastRow = sh.Cells(sh.Rows.Count, REGISTRYNAMECOLUMN).End(xlUp).Row
    If lastRow < DROPDOWNSTARTROW Then Exit Sub

    this.registryRows = lastRow - DROPDOWNSTARTROW + 1
    this.registryBlock = sh.Range( _
        sh.Cells(DROPDOWNSTARTROW, REGISTRYNAMECOLUMN), _
        sh.Cells(lastRow, REGISTRYLONAMECOLUMN)).Value
End Sub

'@label:registry-item
'@sub-title Read one field of one registry row as text
'@param rowIndex Long. Row position inside the registry block.
'@param fieldIndex Long. REGISTRYNAMEITEM or REGISTRYLONAMEITEM.
'@return String. The stored text, or an empty string when out of range.
Private Function RegistryItem(ByVal rowIndex As Long, ByVal fieldIndex As Long) As String
    EnsureRegistry
    If rowIndex < 1 Or rowIndex > this.registryRows Then Exit Function
    RegistryItem = CStr(this.registryBlock(rowIndex, fieldIndex))
End Function

'@label:registry-index-of
'@sub-title Find the registry row that holds a dropdown name
'@param listName String. Dropdown name to look for.
'@return Long. Row position inside the registry block, or 0 when absent.
Private Function RegistryIndexOf(ByVal listName As String) As Long
    Dim rowIndex As Long

    EnsureRegistry

    For rowIndex = 1 To this.registryRows
        If this.registryBlock(rowIndex, REGISTRYNAMEITEM) = listName Then
            RegistryIndexOf = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

'@label:registry-row
'@sub-title Worksheet row of a registry block position
'@param rowIndex Long. Row position inside the registry block.
'@return Long. The 1-based worksheet row.
Private Function RegistryRow(ByVal rowIndex As Long) As Long
    RegistryRow = DROPDOWNSTARTROW + rowIndex - 1
End Function

'@label:registry-free-row
'@sub-title First worksheet row free for a new registry entry
'@details
'A removed dropdown leaves an empty row behind, and that row is used
'again before the registry grows.
'@return Long. The 1-based worksheet row to write on.
Private Function RegistryFreeRow() As Long
    Dim rowIndex As Long

    EnsureRegistry

    For rowIndex = 1 To this.registryRows
        If LenB(RegistryItem(rowIndex, REGISTRYNAMEITEM)) = 0 Then
            RegistryFreeRow = RegistryRow(rowIndex)
            Exit Function
        End If
    Next rowIndex

    RegistryFreeRow = DROPDOWNSTARTROW + this.registryRows
End Function

'@label:has-dropdown
'@sub-title Whether a dropdown name is in the registry held in memory
'@param listName String. Dropdown name to look for.
'@return Boolean. True when the name is registered.
Private Function HasDropdown(ByVal listName As String) As Boolean
    HasDropdown = (RegistryIndexOf(listName) > 0)
End Function

'@label:missing-dropdown
'@sub-title Report a dropdown that is not on the worksheet
'@details
'Logs the given message and prints the debug trace when the dropdown is
'absent. Callers use the return value to leave at once.
'@param listName String. Dropdown name to look for.
'@param message String. Text to log when the dropdown is missing.
'@param showTrace Optional Boolean. When True, prints the debug trace. Defaults to True.
'@return Boolean. True when the dropdown does not exist.
Private Function MissingDropdown(ByVal listName As String, ByVal message As String, _
    Optional ByVal showTrace As Boolean = True) As Boolean

    If HasDropdown(listName) Then Exit Function

    LogInfo message
    If showTrace Then ShowDebug
    MissingDropdown = True
End Function


'@section Counters
'===============================================================================

'@label:internal-store
'@prop-title Lazy-loaded HiddenNames store for workbook or sheet scope
'@details
'Returns the cached HiddenNames instance for either the workbook
'(default) or the worksheet. Creates the store on first access.
'@param sheetLevel Optional Boolean. When True, returns the worksheet-level store. Defaults to False.
'@return HiddenNames. The requested store.
Private Property Get InternalStore(Optional ByVal sheetLevel As Boolean = False) As HiddenNames
    Dim targetSheet As Worksheet
    Dim targetWb As Workbook

    If sheetLevel Then
        If this.shStored Is Nothing Then
            Set targetSheet = Wksh
            If targetSheet Is Nothing Then
                ThrowError ProjectError.ObjectNotInitialized, "DropdownLists requires a worksheet before accessing sheet storage."
            End If
            Set this.shStored = HiddenNames.Create(targetSheet)
        End If
        Set InternalStore = this.shStored
    Else
        If this.wbStored Is Nothing Then
            Set targetSheet = Wksh
            If targetSheet Is Nothing Then
                ThrowError ProjectError.ObjectNotInitialized, "DropdownLists requires a worksheet before accessing workbook storage."
            End If
            Set targetWb = targetSheet.Parent
            Set this.wbStored = HiddenNames.Create(targetWb)
        End If
        Set InternalStore = this.wbStored
    End If
End Property

'@label:counter-name
'@sub-title Stored name of a dropdown counter
'@param sheetLevel Boolean. True for the worksheet counter, False for the workbook one.
'@return String. The hidden name holding the counter.
Private Function CounterName(ByVal sheetLevel As Boolean) As String
    CounterName = IIf(sheetLevel, WORKSHEETCOUNTERNAME, WORKBOOKCOUNTERNAME)
End Function

'@label:counter-start
'@sub-title Start value of a dropdown counter
'@param sheetLevel Boolean. True for the worksheet counter, False for the workbook one.
'@return Long. The value the counter takes when it is created.
Private Function CounterStart(ByVal sheetLevel As Boolean) As Long
    CounterStart = IIf(sheetLevel, WORKSHEETCOUNTERSTART, WORKBOOKCOUNTERSTART)
End Function

'@label:counter-store
'@sub-title Store holding a dropdown counter, created when missing
'@details
'The counters are shared by every DropdownLists object of the workbook,
'so they live in HiddenNames and never in memory. Only the store object
'is cached.
'@param sheetLevel Boolean. True for the worksheet counter, False for the workbook one.
'@return HiddenNames. The store that holds the counter.
Private Function CounterStore(ByVal sheetLevel As Boolean) As HiddenNames
    Dim store As HiddenNames

    Set store = InternalStore(sheetLevel)
    If Not store.HasName(CounterName(sheetLevel)) Then
        store.EnsureName CounterName(sheetLevel), CounterStart(sheetLevel), HiddenNameTypeLong
    End If

    Set CounterStore = store
End Function

'@label:counter-value
'@sub-title Read a dropdown counter
'@param sheetLevel Boolean. True for the worksheet counter, False for the workbook one.
'@return Long. The counter value.
Private Function CounterValue(ByVal sheetLevel As Boolean) As Long
    CounterValue = CounterStore(sheetLevel).ValueAsLong(CounterName(sheetLevel), CounterStart(sheetLevel))
End Function

'@label:step-counter
'@sub-title Move one dropdown counter and write it through at once
'@details
'The worksheet counter never goes below zero. The value is written back
'to the store on every step because other DropdownLists objects read the
'same counter.
'@param sheetLevel Boolean. True for the worksheet counter, False for the workbook one.
'@param stepValue Long. Amount to add, negative to go down.
Private Sub StepCounter(ByVal sheetLevel As Boolean, ByVal stepValue As Long)
    Dim store As HiddenNames
    Dim nextValue As Long

    Set store = CounterStore(sheetLevel)
    nextValue = store.ValueAsLong(CounterName(sheetLevel), CounterStart(sheetLevel)) + stepValue

    If sheetLevel Then
        If nextValue < 0 Then nextValue = 0
    End If

    store.SetValue CounterName(sheetLevel), nextValue
End Sub

'@label:step-counters
'@sub-title Move the worksheet and workbook counters together
'@param stepValue Long. Amount to add, negative to go down.
Private Sub StepCounters(ByVal stepValue As Long)
    StepCounter True, stepValue
    StepCounter False, stepValue
End Sub

'@label:increment-counters
'@sub-title Increment both workbook and worksheet counters
Private Sub IncrementCounters()
    StepCounters CLng(DROPDOWNCOUNTERINCREMENT)
End Sub

'@label:decrement-counters
'@sub-title Decrement both workbook and worksheet counters
Private Sub DecrementCounters()
    StepCounters -CLng(DROPDOWNCOUNTERINCREMENT)
End Sub

'@section Helpers
'===============================================================================

'@label:remove-gridlines
'@sub-title Remove gridlines from a worksheet view
'@param sh Worksheet. Target worksheet.
'@param DisplayZeros Optional Boolean. When True, displays zeros. Defaults to False.
Private Sub RemoveGridLines(ByVal sh As Worksheet, Optional DisplayZeros As Boolean = False)
    Dim View As WorksheetView
    For Each View In sh.Parent.Windows(1).SheetViews
        If View.Sheet.Name = sh.Name Then
            View.DisplayGridlines = False
            View.DisplayZeros = DisplayZeros
            Exit Sub
        End If
    Next
End Sub


'@label:last-column
'@sub-title Find the next available column for a new dropdown
'@details
'Returns the column index two positions to the right of the rightmost
'occupied header cell. This leaves a blank separator column between
'adjacent ListObjects.
'@return Long. The 1-based column index.
Private Function LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()

    With sh
        colIndex = .Cells(DROPDOWNSTARTROW, .Columns.Count).End(xlToLeft).Column
    End With

    LastColumn = colIndex + 2
End Function


'@label:sanitized-list-name
'@sub-title Sanitise a dropdown name for use in ListObject naming
'@param listName String. The dropdown name.
'@return String. The trimmed and space-replaced name.
Private Function SanitizedListName(ByVal listName As String) As String
    SanitizedListName = Replace(Application.WorksheetFunction.Trim(listName), " ", "_")
End Function

'@label:list-object-header-name
'@sub-title Build the ListObject header name for a dropdown
'@param listName String. The dropdown name.
'@return String. The prefixed header name.
Private Function ListObjectHeaderName(ByVal listName As String) As String
    ListObjectHeaderName = headerPrefix() & SanitizedListName(listName)
End Function

'@label:corresponding-workbook-name
'@prop-title Workbook-level name used for the dropdown reference
'@details
'Combines the header prefix with the ListObject header name. When no
'prefix is set, falls back to the DEFAULTNAMEPREFIX constant.
'@param listName String. The dropdown name.
'@return String. The workbook-level reference name.
Private Property Get CorrespondingWorkbookName(ByVal listName As String) As String
    Dim headerName As String

    headerName = ListObjectHeaderName(listName)
    CorrespondingWorkbookName = IIf(LenB(headerPrefix()) = 0, DEFAULTNAMEPREFIX & headerName, headerName)
End Property

'@label:list-object-name
'@prop-title ListObject name for a dropdown
'@details
'Returns the stored ListObject name for an existing dropdown, or
'generates and registers a new one for a dropdown being created. The
'workbook counter is what keeps the generated name unique across the
'whole workbook, so two worksheets can hold a dropdown of the same name.
'@param listName String. Dropdown name.
'@return String. The ListObject name.
Private Property Get ListObjectName(ByVal listName As String) As String
    Dim rowIndex As Long
    Dim loName As String

    rowIndex = RegistryIndexOf(listName)

    If rowIndex = 0 Then
        loName = "list_" & SanitizedListName(listName) & CounterValue(False)
        AddLoName listName, loName
    Else
        loName = RegistryItem(rowIndex, REGISTRYLONAMEITEM)
    End If

    ListObjectName = loName
End Property

'@label:dropdown-list-object
'@sub-title ListObject that holds the values of a dropdown
'@param listName String. The dropdown name.
'@return ListObject. The table on the dropdown worksheet.
Private Function DropdownListObject(ByVal listName As String) As ListObject
    Set DropdownListObject = Wksh().ListObjects(ListObjectName(listName))
End Function

'@label:dropdown-label-cell
'@sub-title Cell that carries the label of a dropdown
'@details
'The label sits one row above the header row of the ListObject.
'@param listName String. The dropdown name.
'@return Range. The label cell.
Private Function DropdownLabelCell(ByVal listName As String) As Range
    Set DropdownLabelCell = DropdownListObject(listName).Range.Cells(0, 1)
End Function

'@label:format-label-cell
'@sub-title Write and format the label cell of a dropdown
'@details
'Used when a dropdown is added, and again after a hyperlink is placed on
'the label, because adding a hyperlink wipes the text and the format.
'@param labelCell Range. The cell above the ListObject header.
'@param labelText String. Text to write. An empty text leaves the value alone.
'@param interiorColor Long. Background colour of the cell.
'@param fontColor Long. Colour of the text.
Private Sub FormatLabelCell(ByVal labelCell As Range, ByVal labelText As String, _
    ByVal interiorColor As Long, ByVal fontColor As Long)

    With labelCell
        If LenB(labelText) > 0 Then .Value = labelText
        .Font.UnderLine = xlUnderlineStyleNone
        .Interior.Color = interiorColor
        .Font.Color = fontColor
        .Font.Bold = True
        .Font.Italic = True
    End With
End Sub

'@label:add-lo-name
'@sub-title Register a dropdown-to-ListObject name mapping
'@param listName String. The dropdown name.
'@param loName String. The corresponding ListObject name.
Private Sub AddLoName(ByVal listName As String, ByVal loName As String)
    Dim sh As Worksheet
    Dim targetRow As Long

    Set sh = Wksh()
    targetRow = RegistryFreeRow()

    sh.Cells(targetRow, REGISTRYNAMECOLUMN).Value = listName
    sh.Cells(targetRow, REGISTRYLONAMECOLUMN).Value = loName

    'The worksheet no longer matches the copy held in memory
    this.registryLoaded = False
End Sub


'@label:remove-lo-name
'@sub-title Remove a dropdown-to-ListObject name mapping
'@param lst String. The dropdown name to unregister.
Private Sub RemoveLoName(ByVal lst As String)
    Dim sh As Worksheet
    Dim rowIndex As Long
    Dim targetRow As Long

    rowIndex = RegistryIndexOf(lst)
    If rowIndex = 0 Then Exit Sub

    Set sh = Wksh()
    targetRow = RegistryRow(rowIndex)
    sh.Range(sh.Cells(targetRow, REGISTRYNAMECOLUMN), _
             sh.Cells(targetRow, REGISTRYLONAMECOLUMN)).ClearContents

    'The worksheet no longer matches the copy held in memory
    this.registryLoaded = False
End Sub


'@label:length
'@prop-title Total number of dropdowns on the worksheet
'@return Long. The dropdown count.
Public Property Get Length() As Long
    Length = CounterValue(True)
End Property


'@label:current-dropdown-label
'@sub-title Build the display label for the next dropdown
'@details
'Combines the prefix with the current counter to produce a
'human-readable label (e.g. "Custom Dropdown 3").
'@param labelPrefix String. Text prefix for the label.
'@return String. The assembled label.
Private Function CurrentDropdownLabel(ByVal labelPrefix As String) As String
    Dim dropCounter As Long

    dropCounter = CounterValue(True)
    CurrentDropdownLabel = labelPrefix & " " & dropCounter
End Function

'@label:exists
'@sub-title Check whether a dropdown exists on the worksheet
'@details
'Looks only at the registry of this worksheet, so two worksheets can
'hold a dropdown of the same name without seeing each other.
'@param listName String. Dropdown name to search for.
'@return Boolean. True when the dropdown is found.
Public Function Exists(ByVal listName As String) As Boolean
    RefreshRegistry
    Exists = HasDropdown(listName)
End Function

'@label:workbook-name
'@prop-title Workbook-level name a dropdown answers to
'@details
'The name SetValidation points a cell at, and the name any other reader uses to
'reach the values. It is derived from the list name and the header prefix, so
'it can be asked for before the dropdown is added.
'
'This is here because a reader outside the class needs the name in a place
'where building a manager is too costly - a worksheet function called once per
'formula cell reads the range by this name and nothing else.
'@param listName String. Dropdown name.
'@return String. The workbook-level name.
Public Property Get WorkbookName(ByVal listName As String) As String
    WorkbookName = CorrespondingWorkbookName(listName)
End Property

'@label:label-range
'@prop-title Label cell above a dropdown ListObject
'@details
'Returns the cell directly above the header row of the specified
'dropdown ListObject. Returns Nothing when the dropdown does not exist.
'@param listName String. Dropdown name to look up.
'@return Range. The label cell, or Nothing.
Public Property Get LabelRange(ByVal listName As String) As Range
    RefreshRegistry
    If Not HasDropdown(listName) Then Exit Property

    'Top header of the range of the listObject
    Set LabelRange = DropdownLabelCell(listName)
End Property


'@label:values
'@prop-title Retrieve the values of a dropdown
'@details
'Returns a BetterArray of the entries stored in the specified dropdown.
'Logs a diagnostic note and exits when the dropdown does not exist.
'@param listName String. Dropdown name to read from.
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@return BetterArray. The dropdown entries.
Public Property Get Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray

    Dim lstVal As BetterArray
    Dim sh As Worksheet
    Dim rng As Range

    Set lstVal = New BetterArray
    lstVal.LowerBound = 1

    RefreshRegistry

    If Not MissingDropdown(listName, _
        "Can not retrive values of a non existing dropdown: " & listName & " not found.") Then

        Set sh = Wksh()
        Set rng = DropdownListObject(listName).Range

        If Not includeHeaders Then
            Set rng = sh.Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, 1))
        End If

        lstVal.FromExcelRange rng, DetectLastColumn:=False, DetectLastRow:=False
    End If

    Set Values = lstVal.Clone()
End Property

'@section Modify
'===============================================================================

'@label:add
'@sub-title Add a new dropdown to the worksheet
'@details
'Creates a one-column ListObject populated with the supplied values,
'registers the name-to-ListObject mapping, and creates a workbook-level
'named reference. Optionally writes a visible label above the ListObject
'and applies a table style. Silently exits when the dropdown already exists.
'The ListObject name is built before the counters move, so it carries the
'old workbook number, while the label is built after and carries the new
'worksheet number.
'@param listValues BetterArray. String values to populate the ListObject.
'@param listName String. Name to assign to the dropdown.
'@param addLabel Optional Boolean. When True, writes a label above the ListObject. Defaults to False.
'@param counterPrefix Optional String. Prefix for the label text. Defaults to vbNullString.
'@param tableStyle Optional String. Excel table style name. Defaults to "TableStyleLight13".
'@export
Public Sub Add(ByVal listValues As BetterArray, ByVal listName As String, _
    Optional ByVal addLabel As Boolean = False, _
    Optional ByVal counterPrefix As String = vbNullString, _
    Optional ByVal tableStyle As String = "TableStyleLight13")


    Dim sh As Worksheet
    Dim colIndex As Long
    Dim headerName As String              'header Name
    Dim loName As String                  'List object name
    Dim rng As Range
    Dim store As HiddenNames
    Dim dropdownLo As ListObject

    RefreshRegistry

    'Do not add a new dropdown if there is already another one
    If HasDropdown(listName) Then
        LogInfo "Dropdown " & listName & " already exits: not added"
        Exit Sub
    End If

    If (listName = LISTNAMECOLUMNHEADER) Then
        LogInfo "Dropdown " & listName & " is an invalid name already used for programming purposes, please change it."
        Exit Sub
    End If

    loName = ListObjectName(listName)
    headerName = ListObjectHeaderName(listName)

    colIndex = LastColumn()
    Set sh = Wksh()

    'Add an increment to the dropdown counter
    IncrementCounters


    'dropdown startRow is where to add the ListObject
    With sh
        listValues.ToExcelRange .Cells(DROPDOWNSTARTROW + 1, colIndex)
        .Cells(DROPDOWNSTARTROW, colIndex).Value = headerName
        Set rng = .Range(.Cells(DROPDOWNSTARTROW, colIndex), .Cells(DROPDOWNSTARTROW + listValues.Length, colIndex))

        'Add A list object table
        '@Ignore ReadOnlyPropertyAssignment
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName

        Set dropdownLo = .ListObjects(loName)
    End With

    'Add the label, with a prefix if needed
    If addLabel Then
        FormatLabelCell sh.Cells(DROPDOWNSTARTROW - 1, colIndex), _
                        CurrentDropdownLabel(counterPrefix), _
                        vbWhite, RGB(91, 155, 213)
    End If

    'Add a style to the listObject
    On Error Resume Next
    dropdownLo.TableStyle = tableStyle
    On Error GoTo 0


    'Add the corresponding name at workbook level
    Set store = InternalStore(False)
    store.SetListObjectHeader CorrespondingWorkbookName(listName), dropdownLo, headerName

    'Auto fit dropdown column, and increment the dropdown counter
    rng.EntireColumn.AutoFit
End Sub

'@label:remove
'@sub-title Remove a dropdown from the worksheet
'@details
'Deletes the ListObject, its workbook-level name reference, the
'internal registry entry, and decrements both counters. Silently
'exits when the dropdown does not exist.
'@param listName String. Dropdown name to remove.
'@export
Public Sub Remove(ByVal listName As String)
    Dim rng As Range
    Dim store As HiddenNames

    RefreshRegistry

    If MissingDropdown(listName, _
        "Can not delete a dropdown that does not exits: " & listName & "Not found", _
        showTrace:=False) Then Exit Sub

    Set store = InternalStore(False)
    store.RemoveName CorrespondingWorkbookName(listName)

    'Remove the listObject
    Set rng = DropdownListObject(listName).Range

    'Important to delete the blank column on the right before the
    'dropdown column itself
    rng.Offset(, 1).EntireColumn.Delete
    rng.EntireColumn.Delete

    'Remove the name from the stored name in the worksheet
    RemoveLoName listName

    'Decrement counters
    DecrementCounters
End Sub

'@label:clear-list
'@sub-title Clear entries of a dropdown
'@details
'Removes all data rows from the dropdown ListObject without deleting
'the dropdown itself. Silently exits when the dropdown does not exist.
'@param listName String. Dropdown name to clear.
'@export
Public Sub ClearList(ByVal listName As String)

    RefreshRegistry

    If MissingDropdown(listName, _
        "Can not clear a dropdown that does not exits: " & listName & " Not found") Then Exit Sub

    'The table may hold no data row at all, so let Excel fail quietly here
    On Error Resume Next
    DropdownListObject(listName).DataBodyRange.Delete
    On Error GoTo 0
End Sub

'@label:update
'@sub-title Update a dropdown with new values
'@details
'Replaces or appends values in the specified dropdown. When
'pasteAtBottom is True, appends below existing data. Duplicate
'entries are removed when removeDuplicates is True. Silently exits
'when the dropdown does not exist or the data is empty.
'@param UpdateData BetterArray. New values to write.
'@param listName String. Dropdown name to update.
'@param pasteAtBottom Optional Boolean. When True, appends below existing rows. Defaults to False.
'@param removeDuplicates Optional Boolean. When True, removes duplicate entries. Defaults to True.
'@export
Public Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _
    Optional ByVal pasteAtBottom As Boolean = False, _
    Optional ByVal removeDuplicates As Boolean = True)

    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim loRng As Range
    Dim Lo As listObject
    Dim tabLength As Long

    RefreshRegistry

    If MissingDropdown(listName, _
        "Can not Update a dropdown that does not exists: " & listName & " Not found") Then Exit Sub

    If  (UpdateData.Length = 0) Then
        LogInfo "Can not Update because updated data is empty"
        Exit Sub
    End If

    Set sh = Wksh()
    Set Lo = DropdownListObject(listName)
    Set loRng = Lo.Range

    If (Not pasteAtBottom) Then
        On Error Resume Next
        Lo.DataBodyRange.ClearContents
        On Error GoTo 0
        Set cellRng = loRng.Cells(2, 1)
        tabLength = UpdateData.Length + 1
    Else
        tabLength = loRng.Rows.Count + UpdateData.Length
        Set cellRng = loRng.Cells(loRng.Rows.Count + 1, 1)
    End If

    'put the new data in the range
    UpdateData.ToExcelRange cellRng

    'Resize the listObject to the new Range
    Set rng = sh.Range(loRng.Cells(1, 1), loRng.Cells(tabLength, 1))
    Lo.Resize rng

    'Remove duplicates if pasteAtBottom (the loRng has changed with a resize)
    If pasteAtBottom Or removeDuplicates Then
        On Error Resume Next
        Lo.Range.RemoveDuplicates Columns:=1, Header:=xlYes
        On Error GoTo 0
    End If
End Sub

'@label:sort
'@sub-title Sort a dropdown ListObject
'@details
'Sorts the dropdown entries using the specified sort order. Silently
'exits when the dropdown does not exist.
'@param listName String. Dropdown name to sort.
'@param sortOrder Optional Long. xlSortOrder enumeration value. Defaults to xlAscending.
'@export
Public Sub Sort(ByVal listName As String, _
    Optional ByVal sortOrder As Long = xlAscending)

    Dim rng As Range

    RefreshRegistry

    If MissingDropdown(listName, _
        "Can not sort a dropdown that does not exits: " & listName & " not found.") Then Exit Sub

    Set rng = DropdownListObject(listName).Range

    On Error Resume Next
    rng.Sort key1:=rng, Order1:=sortOrder, Header:=xlYes
    On Error GoTo 0
End Sub

'@label:translate
'@sub-title Translate all dropdowns on the worksheet
'@details
'Iterates over every registered ListObject and applies the supplied
'TranslationObject to its data body. Skips translation when the
'current language matches and force is False. ListObjects whose name
'contains the exclude substring are skipped.
'@param trads Object. A TranslationObject providing translations.
'@param force Optional Boolean. When True, translates regardless of language match. Defaults to False.
'@param exclude Optional String. Substring filter to skip matching ListObjects.
'@export
Public Sub Translate(ByVal trads As Object, Optional ByVal force As Boolean = False, Optional ByVal exclude As String)
    Dim sh As Worksheet
    Dim loName As String
    Dim Lo As ListObject
    Dim store As HiddenNames
    Dim rowIndex As Long
    Dim skipList As Boolean

    If trads Is Nothing Then Exit Sub
    If TypeName(trads) <> "TranslationObject" Then Exit Sub

    'check language before moving to translation
    Set store = InternalStore()

    If (trads.Language = store.ValueAsString(WORKSHEETLANGUAGENAME)) And (Not force) Then
        Exit Sub
    End If

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    RefreshRegistry

    For rowIndex = this.registryRows To REGISTRYFIRSTDATAITEM Step -1

        loName = RegistryItem(rowIndex, REGISTRYLONAMEITEM)
        skipList = (LenB(loName) = 0)

        If (Not skipList) And (LenB(exclude) > 0) Then
            skipList = (InStr(1, loName, exclude, vbTextCompare) > 0)
        End If

        If Not skipList Then
            Set Lo = Nothing
            'A registered name can point to a table that is already gone
            On Error Resume Next
                Set Lo = sh.ListObjects(loName)
            On Error GoTo 0

            If Not (Lo Is Nothing) Then trads.TranslateRange Lo.DataBodyRange
        End If
    Next rowIndex

    'update language value
    store.EnsureName WORKSHEETLANGUAGENAME, trads.Language, HiddenNameTypeString
    store.SetValue WORKSHEETLANGUAGENAME, trads.Language
End Sub

'@section Operations
'===============================================================================

'@label:set-validation
'@sub-title Apply dropdown validation to a Range
'@details
'Sets data validation on the target Range using the specified dropdown
'as the validation source. alertType controls the severity: "info"
'shows an information alert, "warning" shows a warning, and "error"
'prevents entry of values outside the list. Silently exits when the
'dropdown or range does not exist.
'@param cellRng Range. Target range to validate.
'@param listName String. Dropdown name to use as the validation source.
'@param alertType Optional String. Alert severity ("info", "warning", "error"). Defaults to "info".
'@param message Optional String. Message shown on validation failure. Defaults to vbNullString.
'@param showError Optional Boolean. When True, displays validation feedback. Defaults to True.
'@export
Public Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
    Optional ByVal alertType As String = "info", _
    Optional ByVal message As String = vbNullString, _
    Optional ByVal showError As Boolean = True)

    Dim dropdownName As String

    'This reads the registry and never writes it, so the copy held in memory
    'is good enough. Every routine that changes the registry drops that copy,
    'so a run of validations after an Add still sees the new entry.
    EnsureRegistry

    If MissingDropdown(listName, "There is no available dropdown for " & listName) Then Exit Sub

    If cellRng Is Nothing Then
        LogInfo "There is no range to set the validation of " & listName & " on."
        ShowDebug
        Exit Sub
    End If

    'dropdown names can change. I add the tag
    dropdownName = "= " & CorrespondingWorkbookName(listName)

    With cellRng.validation
        .Delete
        Select Case alertType
            Case "error"
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween, Formula1:=dropdownName
            Case "warning"
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, _
                    Operator:=xlBetween, Formula1:=dropdownName
            Case Else
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                    Operator:=xlBetween, Formula1:=dropdownName
        End Select

        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .errorTitle = vbNullString
        .InputMessage = vbNullString
        .errorMessage = Left(message, 254)
        .ShowInput = showError
        .ShowError = showError
    End With
End Sub


'@label:add-link
'@sub-title Add a hyperlink between a dropdown and a cell
'@details
'Creates a hyperlink that navigates between the dropdown label cell
'and the specified target range. When returnLink is True, the link
'is placed on the dropdown label pointing at targetRng; when False,
'the link is placed on targetRng pointing at the dropdown label.
'Preserves the original cell formatting after link insertion.
'@param listName String. Dropdown name whose label is involved.
'@param targetRng Range. The other end of the hyperlink.
'@param returnLink Optional Boolean. When True, link is on the dropdown side. Defaults to False.
Private Sub AddLink(ByVal listName As String, ByVal targetRng As Range, _
    Optional ByVal returnLink As Boolean = False)

    Dim sh As Worksheet
    Dim anchorRng As Range
    Dim headerRng As Range
    Dim targetAddress As String
    Dim savedValue As String
    Dim savedInteriorColor As Long
    Dim savedFontColor As Long

    'Reads the registry to find the label cell and never writes it, so the
    'copy held in memory is used.
    EnsureRegistry

    If MissingDropdown(listName, _
        "Can not add a returnlink on a non existing dropdown: " & listName & " not found.") Then Exit Sub

    If (targetRng Is Nothing) Then
        LogInfo "Can not add a returnlink to a non existing Range: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    Set headerRng = DropdownLabelCell(listName)

    If returnLink Then
        Set sh =  Wksh()
        Set anchorRng = headerRng
        targetAddress = "'" & targetRng.Parent.Name & "'" & "!" & targetRng.Address
    Else
        Set sh = targetRng.Parent
        Set anchorRng = targetRng
        targetAddress = "'" & headerRng.Parent.Name & "'" & "!" & headerRng.Address
    End If


    savedValue = anchorRng.Value
    savedInteriorColor = anchorRng.Interior.Color
    savedFontColor = anchorRng.Font.Color

    'Adding hyperlinks remove the writings
    '(sub label and so on, so need to set it back)
    With sh
        .Hyperlinks.Add _
            Anchor:=anchorRng, _
            Address:="", _
            SubAddress:= targetAddress, _
            TextToDisplay:= ""
    End With

    'Adding links remove the formatting, set it back
    FormatLabelCell anchorRng, savedValue, savedInteriorColor, savedFontColor
End Sub

'@label:add-return-link
'@sub-title Add a return hyperlink on the dropdown header
'@details
'Adds a hyperlink on the dropdown label cell that navigates back to
'the specified range.
'@param listName String. Dropdown whose label receives the link.
'@param returnRng Range. Target destination of the hyperlink.
'@export
Public Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    AddLink listName:=listName, targetRng:=returnRng, returnLink:=True
End Sub


'@label:add-hyperlink
'@sub-title Add a hyperlink on a cell targeting a dropdown
'@details
'Adds a hyperlink on the specified range that navigates to the label
'cell of the dropdown.
'@param listName String. Dropdown to navigate to.
'@param linkRng Range. Cell that receives the hyperlink.
'@export
Public Sub AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
    AddLink listName:=listName, targetRng:=linkRng, returnLink:=False
End Sub


'@section Checkings
'===============================================================================

'@label:log-info
'@sub-title Record a diagnostic entry
'@param label String. The message to log.
'@param scope Optional Byte. Severity level. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
    Optional ByVal scope As Byte = checkingNote)

    Dim checkName As String

    checkName = "DropdownList " & Name() & " object trace"

    'First, instanciante the internalCheks Objects
    If internalChecks Is Nothing Then Set internalChecks = Checking.Create(checkName)

    'Add new informations to the checking object
    internalChecks.Add Cstr(checkCounter), label, scope
    checkCounter = checkCounter + 1

    If Not this.hasChecks Then this.hasChecks = True
End Sub

'@label:has-checkings
'@prop-title Whether the dropdown object has logged diagnostic messages
'@return Boolean. True when diagnostic entries exist.
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. An Checking instance, or Nothing.
Public Property Get CheckingValues() As Object

    'If there is no checking, exit the property
    If Not HasCheckings() Then  Exit Property

    Set CheckingValues = internalChecks
End Property

'@label:show-debug
'@sub-title Print a debug trace to the Immediate window
Private Sub ShowDebug()
    Debug.Print "Check the informations of dropdownlist object " & Name()
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumb Long. ProjectError code.
'@param errorMessage String. Descriptive message.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumb As Long, ByVal errorMessage As String)
    Err.Raise errNumb, CLASSNAME, "Error : " & Chr(34) & errorMessage
End Sub

'@label:seal
'@sub-title Seal the instance so setup setters can no longer be written
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Guard a setup setter against writes after sealing
'@param propName String. Name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub
