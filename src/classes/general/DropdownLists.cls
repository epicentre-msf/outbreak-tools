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
Implements IDropdownLists

'@class DropdownLists
'@description
'Manages worksheet-based dropdown lists backed by one-column ListObjects.
'Each dropdown is registered with an internal name-to-ListObject mapping
'and a workbook-level named reference. Supports adding, removing,
'updating, sorting, translating, and validating dropdown entries, plus
'hyperlink navigation between the dropdown sheet and consuming cells.
'@depends HiddenNames, BetterArray, Checking
'@version 1.0 (2026-02-09)
'@author OutbreakTools

Private Type TDropdownLists
    sh As Worksheet
    hprefix As String
    hasChecks As Boolean
    initialised As Boolean
    wbStored As IHiddenNames
    shStored As IHiddenNames
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


'checking objects
Private checkCounter As Long
Private internalChecks As IChecking

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
'@return IDropdownLists. The new instance.
'@throws ProjectError.ElementNotFound When outsh is Nothing.
'@export
Public Function Create(ByVal outsh As Worksheet, Optional ByVal hprefix As String = "dropdown_") As IDropdownLists

    If outsh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, "Dropdown worksheet not found"
    End If

    With New DropdownLists
        Set .Wksh = outsh
        Let .headerPrefix = hprefix
        Set Create = .Self()
    End With
End Function

'@section Elements
'===============================================================================

'@label:self
'@prop-title Self-reference returning the interface pointer
'@return IDropdownLists. This instance.
Public Property Get Self() As IDropdownLists
    Set Self = Me
End Property

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
    
    Dim store As IHiddenNames
    Set store = HiddenNames.Create(outsh)
    this.initialised = (store.ValueAsString(DROPDOWNINITIALSEDNAME) = "Yes")

    If (Not this.initialised) Then

       'Format the worksheet before creating dropdowns
        RemoveGridLines outsh

        outsh.Range("A:D").Font.Color = vbWhite
        outsh.Range("A:D").Interior.Color = vbWhite

        outsh.Cells(DROPDOWNSTARTROW, 3).Value = LISTNAMECOLUMNHEADER
        outsh.Cells(DROPDOWNSTARTROW, 4).Value = LONAMECOLUMNHEADER

        this.initialised = True
        store.EnsureName DROPDOWNINITIALSEDNAME, "Yes", HiddenNameTypeString
    End If

    Set this.sh = outsh
    Set this.shStored = Nothing
    Set this.wbStored = Nothing
End Property


'@label:all-dropdowns
'@prop-title All registered dropdown names
'@details
'Scans the internal registry column and collects every dropdown name
'into a BetterArray.
'@return BetterArray. The dropdown name list.
Private Property Get AllDropdowns() As BetterArray
    Dim sh As Worksheet
    Dim temp As BetterArray
    Dim headerCell As Range
    Dim cursor As Range

    Set temp = New BetterArray
    temp.LowerBound = 1

    Set sh = Wksh()
    If sh Is Nothing Then
        Set AllDropdowns = temp
        Exit Property
    End If

    Set headerCell = sh.Cells(DROPDOWNSTARTROW, 3)
    Set cursor = sh.Cells(sh.Rows.Count, 3).End(xlUp)

    If cursor.Row <= headerCell.Row Then
        Set AllDropdowns = temp
        Exit Property
    End If

    Do While cursor.Row > headerCell.Row
        If LenB(Trim$(CStr(cursor.Value))) > 0 Then
            temp.Push CStr(cursor.Value)
        End If
        Set cursor = cursor.Offset(-1)
    Loop

    temp.Reverse
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
    this.hprefix = hprefix
End Property

'@label:name
'@prop-title Name of the dropdown object
'@details
'Returns the name of the worksheet that hosts the dropdown lists.
'@return String. The worksheet name.
Private Property Get Name() As String
    Name = Wksh.Name
End Property


'@section Workbook Counter
'===============================================================================

'@label:internal-store
'@prop-title Lazy-loaded HiddenNames store for workbook or sheet scope
'@details
'Returns the cached IHiddenNames instance for either the workbook
'(default) or the worksheet. Creates the store on first access.
'@param sheetLevel Optional Boolean. When True, returns the worksheet-level store. Defaults to False.
'@return IHiddenNames. The requested store.
Private Property Get InternalStore(Optional ByVal sheetLevel As Boolean = False) As IHiddenNames
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


'@label:ensure-workbook-counter
'@sub-title Ensure the workbook-level dropdown counter exists
Private Sub EnsureWorkbookCounter()
    Dim store As IHiddenNames

    Set store = InternalStore(False)
    If Not store.HasName(WORKBOOKCOUNTERNAME) Then
        store.EnsureName WORKBOOKCOUNTERNAME, 1, HiddenNameTypeLong
    End If
End Sub

'@label:workbook-counter-value
'@sub-title Read the current workbook dropdown counter
'@return Long. The counter value.
Private Function WorkbookCounterValue() As Long
    Dim store As IHiddenNames

    EnsureWorkbookCounter
    Set store = InternalStore(False)
    WorkbookCounterValue = store.ValueAsLong(WORKBOOKCOUNTERNAME, 1)
End Function

'@label:set-workbook-counter
'@sub-title Write a new workbook dropdown counter value
'@param newValue Long. The value to store.
Private Sub SetWorkbookCounter(ByVal newValue As Long)
    Dim store As IHiddenNames

    EnsureWorkbookCounter
    Set store = InternalStore(False)
    store.SetValue WORKBOOKCOUNTERNAME, newValue
End Sub

'@label:increment-workbook-counter
'@sub-title Increment the workbook dropdown counter by one
Private Sub IncrementWorkbookCounter()
    Dim nextValue As Long
    nextValue = WorkbookCounterValue() + DROPDOWNCOUNTERINCREMENT
    SetWorkbookCounter nextValue
End Sub

'@label:decrement-workbook-counter
'@sub-title Decrement the workbook dropdown counter by one
Private Sub DecrementWorkbookCounter()
     Dim nextValue As Long
    nextValue = WorkbookCounterValue() - DROPDOWNCOUNTERINCREMENT
    SetWorkbookCounter nextValue
End Sub
'@section Worksheet Counter
'===============================================================================

'@label:ensure-worksheet-counter
'@sub-title Ensure the worksheet-level dropdown counter exists
Private Sub EnsureWorksheetCounter()
    Dim store As IHiddenNames

    Set store = InternalStore(True)
    If Not store.HasName(WORKSHEETCOUNTERNAME) Then
        store.EnsureName WORKSHEETCOUNTERNAME, 0, HiddenNameTypeLong
    End If
End Sub


'@label:sheet-counter-value
'@sub-title Read the current worksheet dropdown counter
'@return Long. The counter value.
Private Function SheetCounterValue() As Long
    Dim store As IHiddenNames

    EnsureWorksheetCounter
    Set store = InternalStore(True)
    SheetCounterValue = store.ValueAsLong(WORKSHEETCOUNTERNAME, 0)
End Function

'@label:set-sheet-counter
'@sub-title Write a new worksheet dropdown counter value
'@param newValue Long. The value to store (clamped to 0).
Private Sub SetSheetCounter(ByVal newValue As Long)
    Dim store As IHiddenNames

    EnsureWorksheetCounter
    If newValue < 0 Then newValue = 0

    Set store = InternalStore(True)
    store.SetValue WORKSHEETCOUNTERNAME, newValue
End Sub

'@label:increment-sheet-counter
'@sub-title Increment the worksheet dropdown counter by one
Private Sub IncrementSheetCounter()
    Dim nextValue As Long
    nextValue = SheetCounterValue() + DROPDOWNCOUNTERINCREMENT
    SetSheetCounter nextValue
End Sub

'@label:decrement-sheet-counter
'@sub-title Decrement the worksheet dropdown counter by one
Private Sub DecrementSheetCounter()
    Dim nextValue As Long
    nextValue = SheetCounterValue() - DROPDOWNCOUNTERINCREMENT
    SetSheetCounter nextValue
End Sub

'@label:increment-counters
'@sub-title Increment both workbook and worksheet counters
Private Sub IncrementCounters()
    IncrementSheetCounter
    IncrementWorkbookCounter
End Sub

'@label:decrement-counters
'@sub-title Decrement both workbook and worksheet counters
Private Sub DecrementCounters()
    DecrementSheetCounter
    DecrementWorkbookCounter
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
'Returns the cached ListObject name for an existing dropdown, or
'generates and registers a new one for a dropdown being created.
'@param listName String. Dropdown name.
'@return String. The ListObject name.
Private Property Get ListObjectName(ByVal listName As String) As String
    Dim loName As String

    If Not Exists(listName) Then
        loName = "list_" & SanitizedListName(listName) & WorkbookCounterValue()
        AddLoName listName, loName
    Else
        loName = RetrieveLoName(listName)
    End If

    ListObjectName = loName
End Property

'@label:add-lo-name
'@sub-title Register a dropdown-to-ListObject name mapping
'@param listName String. The dropdown name.
'@param loName String. The corresponding ListObject name.
Private Sub AddLoName(ByVal listName As String, ByVal loName As String)

    Dim cellRng As Range
    Dim sh As Worksheet

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    Do While (LenB(CStr(cellRng.Value)) <> 0)
        Set cellRng = cellRng.Offset(1)
    Loop

    cellRng.Value = listName
    cellRng.Offset(, 1).Value = loName
End Sub


'@label:retrieve-lo-name
'@prop-title Retrieve the ListObject name for a dropdown
'@param lst String. The dropdown name.
'@return String. The ListObject name, or vbNullString when not found.
Private Property Get RetrieveLoName(ByVal lst As String) As String
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim nbElmts As Long

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    For nbElmts = 0 To Length()
        If cellRng.Value = lst Then
            RetrieveLoName = cellRng.Offset(, 1).Value
            Exit Property
        End If
        Set cellRng = cellRng.Offset(1)
    Next
End Property

'@label:remove-lo-name
'@sub-title Remove a dropdown-to-ListObject name mapping
'@param lst String. The dropdown name to unregister.
Private Sub RemoveLoName(ByVal lst As String)
    Dim cellRng As Range
    Dim sh As Worksheet
    Dim nbElmts As Long

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    For nbElmts = 0 To Length()
        If cellRng.Value = lst Then
            cellRng.ClearContents
            cellRng.Offset(, 1).ClearContents
            Exit Sub
        End If
        Set cellRng = cellRng.Offset(1)
    Next
End Sub


'@label:length
'@prop-title Total number of dropdowns on the worksheet
'@return Long. The dropdown count.
Private Property Get Length() As Long
    Length = SheetCounterValue()
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

    dropCounter = SheetCounterValue()
    CurrentDropdownLabel = labelPrefix & " " & dropCounter
End Function

'@label:exists
'@sub-title Check whether a dropdown exists on the worksheet
'@param listName String. Dropdown name to search for.
'@return Boolean. True when the dropdown is found.
Private Function Exists(ByVal listName As String) As Boolean

    Dim sh As Worksheet
    Dim cellRng As Range
    Dim count As Long

    Set sh = Wksh()
    Set cellRng = sh.Cells(DROPDOWNSTARTROW, 3)

    For count = 0 To Length()
        If cellRng.Value = listName Then
            Exists = True
            Exit Function
        End If
        Set cellRng = cellRng.Offset(1)
    Next

End Function

'@label:label-range
'@prop-title Label cell above a dropdown ListObject
'@details
'Returns the cell directly above the header row of the specified
'dropdown ListObject. Returns Nothing when the dropdown does not exist.
'@param listName String. Dropdown name to look up.
'@return Range. The label cell, or Nothing.
Private Property Get LabelRange(ByVal listName As String) As Range

    Dim loName As String
    Dim sh As Worksheet

    If Not Exists(listName) Then Exit Property
    Set sh = Wksh()
    loName = ListObjectName(listName)
    'Top header of the range of the listObject
    Set LabelRange = sh.ListObjects(loName).Range.Cells(0, 1)
End Property


'@label:values
'@prop-title Retrieve the values of a dropdown
'@details
'Returns a BetterArray of the entries stored in the specified dropdown.
'Logs a diagnostic note and exits when the dropdown does not exist.
'@param listName String. Dropdown name to read from.
'@param includeHeaders Optional Boolean. When True, includes the header row. Defaults to False.
'@return BetterArray. The dropdown entries.
Private Property Get Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray

    Dim lstVal As BetterArray
    Dim sh As Worksheet
    Dim rng As Range
    Dim loName As String
    Dim lo As ListObject

    Set lstVal = New BetterArray
    lstVal.LowerBound = 1

    If Not Exists(listName) Then
        LogInfo "Can not retrive values of a non existing dropdown: " & listName & " not found."
        ShowDebug
    Else
        Set sh = Wksh()
        loName = ListObjectName(listName)
        Set Lo = sh.ListObjects(loName)
        Set rng = Lo.Range

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
'@param listValues BetterArray. String values to populate the ListObject.
'@param listName String. Name to assign to the dropdown.
'@param addLabel Optional Boolean. When True, writes a label above the ListObject. Defaults to False.
'@param counterPrefix Optional String. Prefix for the label text. Defaults to vbNullString.
'@param tableStyle Optional String. Excel table style name. Defaults to "TableStyleLight13".
'@export
Private Sub Add(ByVal listValues As BetterArray, ByVal listName As String, _
    Optional ByVal addLabel As Boolean = False, _
    Optional ByVal counterPrefix As String = vbNullString, _
    Optional ByVal tableStyle As String = "TableStyleLight13")


    Dim sh As Worksheet
    Dim colIndex As Long
    Dim headerName As String              'header Name
    Dim loName As String                  'List object name
    Dim rng As Range
    Dim store As IHiddenNames
    Dim dropdownLo As ListObject

    'Do not add a new dropdown if there is already another one
    If Exists(listName) Then
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

        'Add the label, with a prefix if needed
        If addLabel Then
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Value = CurrentDropdownLabel(counterPrefix)
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Interior.Color = vbWhite
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Color = RGB(91, 155, 213)
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Bold = True
            .Cells(DROPDOWNSTARTROW - 1, colIndex).Font.Italic = True
        End If
        
        Set dropdownLo = .ListObjects(loName)
    End With

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
Private Sub Remove(ByVal listName As String)
    Dim sh As Worksheet
    Dim rng As Range
    Dim loName As String
    Dim store As IHiddenNames

    If Not Exists(listName) Then
        LogInfo "Can not delete a dropdown that does not exits: " & listName & "Not found"
        Exit Sub
    End If


    Set store = InternalStore(False)
    store.RemoveName CorrespondingWorkbookName(listName)

    'Remove the listObject
    loName = ListObjectName(listName)
    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range

    'Important to delete the range at the left before deleting the
    'range itself
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
Private Sub ClearList(ByVal listName As String)

    Dim sh As Worksheet
    Dim loName As String

    If Not Exists(listName) Then
        LogInfo "Can not clear a dropdown that does not exits: " & listName & " Not found"
        ShowDebug
        Exit Sub
    End If

    loName = ListObjectName(listName)
    Set sh = Wksh()
    On Error Resume Next
    sh.ListObjects(loName).DataBodyRange.Delete
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
Private Sub Update(ByVal UpdateData As BetterArray, ByVal listName As String, _
    Optional ByVal pasteAtBottom As Boolean = False, _
    Optional ByVal removeDuplicates As Boolean = True)

    Dim sh As Worksheet
    Dim cellRng As Range
    Dim rng As Range
    Dim loRng As Range
    Dim loName As String
    Dim Lo As listObject
    Dim tabLength As Long

    If Not Exists(listName) Then
        LogInfo "Can not Update a dropdown that does not exists: " & listName & " Not found"
        ShowDebug
        Exit Sub
    End If

    If  (UpdateData.Length = 0) Then
        LogInfo "Can not Update because updated data is empty"
        Exit Sub
    End If

    Set sh = Wksh()
    loName = ListObjectName(listName)
    Set Lo = sh.ListObjects(loName)
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
Private Sub Sort(ByVal listName As String, _
    Optional ByVal sortOrder As Long = xlAscending)

    Dim loName As String
    Dim rng As Range
    Dim sh As Worksheet

    If Not Exists(listName) Then
        LogInfo "Can not sort a dropdown that does not exits: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    loName = ListObjectName(listName)
    Set sh = Wksh()
    Set rng = sh.ListObjects(loName).Range

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
Private Sub Translate(ByVal trads As Object, Optional ByVal force As Boolean = False, Optional ByVal exclude As String)
    Dim sh As Worksheet
    Dim headerCell As Range
    Dim cursor As Range
    Dim loName As String
    Dim lo As ListObject
    Dim store As IHiddenNames

    If trads Is Nothing Then Exit Sub
    If TypeName(trads) <> "TranslationObject" Then Exit Sub

    'check language before moving to translation
    Set store = InternalStore()
    
    If (trads.Language = store.ValueAsString(WORKSHEETLANGUAGENAME)) And (Not force) Then
        Exit Sub
    End If

    Set sh = Wksh()
    If sh Is Nothing Then Exit Sub

    Set headerCell = sh.Cells(DROPDOWNSTARTROW, 4)
    Set cursor = sh.Cells(sh.Rows.Count, 4).End(xlUp)

    If cursor.Row <= headerCell.Row Then
        Exit Sub
    End If

    Do While cursor.Row > headerCell.Row
        loName = CStr(cursor.Value) 
        If LenB(loName) = 0 Then GoTo NextLo
        If LenB(exclude) > 0 Then
            If InStr(1, loName, exclude, vbTextCompare) > 0 Then GoTo NextLo
        End If

        On Error Resume Next
            Set lo = sh.ListObjects(loName)
        On Error GoTo 0 
        
        If Not (lo Is Nothing) Then trads.TranslateRange lo.DataBodyRange   
NextLo:    
        Set lo = Nothing
        Set cursor = cursor.Offset(-1)
    Loop

    'update language value
    store.EnsureName WORKSHEETLANGUAGENAME, trads.Language, HiddenNameTypeString
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
Private Sub SetValidation(ByVal cellRng As Range, ByVal listName As String, _
    Optional ByVal alertType As String = "info", _
    Optional ByVal message As String = vbNullString, _
    Optional ByVal showError As Boolean = True)

    Dim dropdownName As String

    If Not Exists(listName) Then
        LogInfo "There is no available dropdown for " & listName
        ShowDebug
        Exit Sub
    End If

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

    Dim loName As String
    Dim sh As Worksheet
    Dim anchorRng As Range
    Dim headerRng As Range
    Dim targetAddress As String
    Dim savedValue As String
    Dim savedInteriorColor As Long
    Dim savedFontColor As Long


    If Not Exists(listName) Then
        LogInfo "Can not add a returnlink on a non existing dropdown: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    If (targetRng Is Nothing) Then
        LogInfo "Can not add a returnlink to a non existing Range: " & listName & " not found."
        ShowDebug
        Exit Sub
    End If

    loName = ListObjectName(listName)
    Set headerRng = Wksh.ListObjects(loName).Range.Cells(0, 1)

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
    With anchorRng
        If savedValue <> vbNullString Then .Value = savedValue
        .Font.UnderLine = xlUnderlineStyleNone
        .Interior.Color = savedInteriorColor
        .Font.Color = savedFontColor
        .Font.Bold = True
        .Font.Italic = True
    End With
End Sub

'@label:add-return-link
'@sub-title Add a return hyperlink on the dropdown header
'@details
'Adds a hyperlink on the dropdown label cell that navigates back to
'the specified range.
'@param listName String. Dropdown whose label receives the link.
'@param returnRng Range. Target destination of the hyperlink.
'@export
Private Sub AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
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
Private Sub AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
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
Private Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title Retrieve the diagnostic log entries
'@return Object. An IChecking instance, or Nothing.
Private Property Get CheckingValues() As Object

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

'@section Interface
'===============================================================================

Private Property Get IDropdownLists_Wksh() As Worksheet
    Set IDropdownLists_Wksh = Wksh()
End Property

Private Property Get IDropdownLists_Exists(ByVal listName As String) As Boolean
    IDropdownLists_Exists = Exists(listName)
End Property

Private Property Get IDropdownLists_Name() As String
    IDropdownLists_Name = Name()
End Property

Private Property Get IDropdownLists_LabelRange(ByVal listName As String) As Range
    Set IDropdownLists_LabelRange = LabelRange(listName)
End Property

Private Property Get IDropdownLists_Values(ByVal listName As String, Optional ByVal includeHeaders As Boolean = False) As BetterArray
    Set IDropdownLists_Values = Values(listName, includeHeaders)
End Property

Private Property Get IDropdownLists_Length() As Long
    IDropdownLists_Length = Length()
End Property

Private Property Get IDropdownLists_AllDropdowns() As BetterArray
    Set IDropdownLists_AllDropdowns = AllDropdowns
End Property

Private Sub IDropdownLists_Add(ByVal listValues As BetterArray, _
    ByVal listName As String, _
    Optional ByVal addLabel As Boolean = False, _
    Optional ByVal counterPrefix As String = vbNullString, _
    Optional ByVal tableStyle As String = "TableStyleLight13")

    Add listValues:=listValues, listName:=listName, _
        addLabel:=addLabel, counterPrefix:=counterPrefix, tableStyle:=tableStyle
End Sub

Private Sub IDropdownLists_Remove(ByVal listName As String)
    Remove listName:=listName
End Sub

Private Sub IDropdownLists_ClearList(ByVal listName As String)
    ClearList listName
End Sub

Private Sub IDropdownLists_Update(ByVal UpdateData As BetterArray, _
    ByVal listName As String, _
    Optional ByVal pasteAtBottom As Boolean = False, _
    Optional ByVal removeDuplicates As Boolean = True)
    Update UpdateData:=UpdateData, _
        listName:=listName, _
        pasteAtBottom:=pasteAtBottom, _
        removeDuplicates:=removeDuplicates
End Sub

Private Sub IDropdownLists_Sort(ByVal listName As String, _
    Optional ByVal sortOrder As Long = xlAscending)
    Sort listName, sortOrder
End Sub

Private Sub IDropdownLists_Translate(ByVal trads As Object, Optional ByVal force As Boolean = False, Optional ByVal exclude As String)
    Translate trads, force
End Sub

Private Sub IDropdownLists_SetValidation(ByVal cellRng As Range, ByVal listName As String, _
    Optional ByVal alertType As String = "info", _
    Optional ByVal message As String = vbNullString, _
    Optional ByVal showError As Boolean = True)

    SetValidation cellRng:=cellRng, listName:=listName, _
        alertType:=alertType, message:=message, _
        showError:=showError
End Sub

Private Sub IDropdownLists_AddReturnLink(ByVal listName As String, ByVal returnRng As Range)
    AddReturnLink listName, returnRng
End Sub


Private Sub IDropdownLists_AddHyperLink(ByVal listName As String, ByVal linkRng As Range)
    AddHyperLink listName, linkRng
End Sub

Private Property Get IDropdownLists_HasCheckings() As Boolean
    IDropdownLists_HasCheckings = HasCheckings()
End Property

Private Property Get IDropdownLists_CheckingValues() As Object
    Set IDropdownLists_CheckingValues = CheckingValues()
End Property
