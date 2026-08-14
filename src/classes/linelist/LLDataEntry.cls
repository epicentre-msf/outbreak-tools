VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLDataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds HList or VList data entry sheets with sections, formatting, and metadata"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Builds HList or VList data entry sheets with sections, formatting, and metadata")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================
'@description
'Data-entry sheet layers built by LLDataEntry. These live on the class
'itself, so depending on LLDataEntry resolves them everywhere.
Public Enum LLDataEntryLayer
    LLDataEntryLayerHList = 1
    LLDataEntryLayerVList = 2
End Enum


'@class LLDataEntry
'@description
'Builds one data entry worksheet of a generated linelist from the dictionary.
'Consumers call Create with a layer mode, a sheet name and the linelist facade,
'then call Build. Entries filed along the way leave through HasCheckings and
'CheckingValues.
'
'WHAT EACH LAYER GETS
'-------------------------------------------------------------------------------
'An HList sheet gets the design format, frozen panes, five action buttons, its
'sections, the GoTo dropdown, the sheet metadata, a ListObject, a printed
'companion sheet, a filtered companion sheet, a pivot table and protection.
'
'A VList sheet gets the design format, a freeze on row 1, its sections, the
'GoTo dropdown, the sheet metadata, a PLAGEVALUES named range and protection.
'The ListObject, the two companion sheets, the pivot table and the action
'buttons belong to the HList layer alone.
'
'WHAT A FAILURE LEAVES BEHIND
'-------------------------------------------------------------------------------
'Both builders protect their worksheet whichever way they leave, file the
'failure for the generation report and raise it again. The caller offers the
'user the incomplete workbook to inspect, so a sheet that failed part way
'through is reachable and it is protected when the user gets there.
'
'Screen updating belongs to the caller. EventsDesignerAdvanced.clickGenerate
'turns it off through ApplicationState for the whole generation and restores it
'on both its exits.
'
'PUBLISHED STATE
'-------------------------------------------------------------------------------
'The worksheet-level hidden names this class writes, and who reads each:
'  sheet_type       EventLinelist, LLExporter, LLImporter, LLGeo, LLSpatial,
'                   ExportButton, CustomLinelistFunctions
'  table_name       EventLinelist, LLImporter, LLGeo, LLSpatial,
'                   EventsLinelistButtons
'  blank_row_count  LLImporter.HasData, EventsLinelistButtons. It holds the
'                   number of filled cells an untouched data row carries, which
'                   is the threshold CustomTable.RemoveRows compares against.
'  filtered_sheet   EventLinelist, LLExporter, LLSpatial, ExportButton
'  <table>_go_to_section   the GoTo dropdown validation
'@depends Linelist, LLSheets, LinelistSpecs, LLdictionary, LLFormat,
'  DropdownLists, Passwords, TranslationObject, Buttons, CustomPivotTable,
'  HiddenNames, SectionBuilder, Checking, BetterArray
'@version 1.1 (2026-07-31)



Private Const CLASS_NAME As String = "LLDataEntry"
Private Const GOTO_SECTION_CODE As String = "go_to_section"
Private Const FILTERED_PREFIX As String = "f"
Private Const PRINT_PREFIX As String = "print_"

'The prefix Linelist.ScopedName puts on a CRF companion. Spelled the same way
'there and in EventsLinelistButtons, and it is part of the sheet's stored table
'name, so it is copied here exactly.
Private Const CRF_PREFIX As String = "crf_"

'What Excel accepts for a worksheet name. Linelist.ScopedName cuts to the same
'figure, so a longer name still builds; this class reads it to tell the user
'which dictionary names come out shortened.
Private Const SHEET_NAME_MAX_LENGTH As Long = 31

'How far the printed companion is grown once its ListObject is made, and the
'same headroom the filtered companion is given.
Private Const COMPANION_SPARE_ROWS As Long = 10

'Button shape codes
Private Const SHP_GEO As String = "SHP_GeoApps"
Private Const SHP_ADD_ROWS As String = "SHP_Add200L"
Private Const SHP_REMOVE_FILT As String = "SHP_ClearFilters"
Private Const SHP_SHOW_HIDE As String = "SHP_ShowHide"
Private Const SHP_OPEN_PRINT As String = "SHP_OpenPrint"

Private Type TLLDataEntry
    layer As Byte
    sheetName As String
    ll As Linelist
    sheetInfo As LLSheets
    internalChecks As Checking
    checkCounter As Long
    hasChecks As Boolean
    builder As SectionBuilder
    sealed As Boolean
End Type

Private this As TLLDataEntry


'@section Instantiation
'===============================================================================
'@description Factory method and Friend bindings for construction.

'@label:create
'@sub-title Create a LLDataEntry instance
'@details
'Creates a new LLDataEntry wired to build a data entry sheet of the given
'layer type (HList or VList) for the named sheet within the linelist. The
'linelist object, the layer, the sheet name and its presence in the dictionary
'are all validated before construction.
'
'The caller may hand in the LLSheets instance it already holds. LLSheets caches
'the resolved dictionary row and the variable count per instance, so one
'instance shared across a whole build searches the dictionary once per sheet.
'Five instances searched it five times.
'@param layer Byte. The LLDataEntryLayer mode (HList = 1, VList = 2).
'@param sheetName String. The dictionary sheet name to build.
'@param ll Linelist. The linelist workbook facade providing worksheets,
'  dropdowns, pivots, and specs.
'@param sheetInfo Optional LLSheets. Sheet metadata over the same dictionary.
'  A new one is built when the caller has none.
'@return LLDataEntry. A fully initialised instance ready for Build.
'@throws ProjectError.ObjectNotInitialized When ll is Nothing.
'@throws ProjectError.InvalidArgument When sheetName is empty, when the layer
'  is unknown, or when the sheet is absent from the dictionary.
'@export
Public Function Create(ByVal layer As Byte, _
                       ByVal sheetName As String, _
                       ByVal ll As Linelist, _
                       Optional ByVal sheetInfo As LLSheets) As LLDataEntry

    Dim llshs As LLSheets
    Dim instance As LLDataEntry

    If ll Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist object is required"
    End If

    If LenB(Trim$(sheetName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheet name cannot be empty"
    End If

    If layer <> LLDataEntryLayerHList And layer <> LLDataEntryLayerVList Then
        ThrowError ProjectError.InvalidArgument, _
                   "There is no data entry layer '" & CStr(layer) & "' in a linelist"
    End If

    Set llshs = sheetInfo
    If llshs Is Nothing Then Set llshs = LLSheets.Create(ll.Dictionary)

    If Not llshs.Contains(sheetName) Then
        ThrowError ProjectError.InvalidArgument, _
                   "The sheet name " & sheetName & " is not present in the dictionary"
    End If

    Set instance = New LLDataEntry

    instance.BinLayer = layer
    instance.BinSheetName = sheetName
    Set instance.BinLinelist = ll
    Set instance.BinSheetInfo = llshs
    instance.Seal
    instance.ReportSheetNameLength

    Set Create = instance
End Function

'@label:seal
'@sub-title Close the instance to creation-only writes
'@details
'Called by the factory once the four bindings are in place. Swapping the layer
'or the sheet name part way through a build would write one sheet's metadata
'onto another, so the setters refuse a second write.
'@export
Public Sub Seal()
    this.sealed = True
End Sub


'@section Bindings
'===============================================================================
'@description Friend setters used only during factory construction.

'@label:bin-layer
'@prop-title Assign the builder layer mode
'@param value Byte. The LLDataEntryLayer constant (HList or VList).
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Let BinLayer(ByVal value As Byte)
    GuardNotSealed "BinLayer"
    this.layer = value
End Property

'@label:bin-sheet-name
'@prop-title Assign the target sheet name
'@param value String. The dictionary sheet name to build.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Let BinSheetName(ByVal value As String)
    GuardNotSealed "BinSheetName"
    this.sheetName = value
End Property

'@label:bin-linelist
'@prop-title Assign the linelist facade
'@param value Linelist. The linelist workbook facade.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Set BinLinelist(ByVal value As Linelist)
    GuardNotSealed "BinLinelist"
    Set this.ll = value
End Property

'@label:bin-sheet-info
'@prop-title Assign the sheet metadata reader
'@param value LLSheets. Sheet metadata over the linelist dictionary.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Set BinSheetInfo(ByVal value As LLSheets)
    GuardNotSealed "BinSheetInfo"
    Set this.sheetInfo = value
End Property

'@label:report-sheet-name-length
'@sub-title File a report entry for a sheet name whose companions come out shortened
'@details
'Excel accepts 31 characters and Linelist.ScopedName cuts every worksheet name
'to that, on creation and on lookup alike, so a long name still builds. What the
'user loses is the choice of what the shortened companion is called, and this
'says so at validation time, before any sheet is made.
'
'The printed companion carries the longest prefix, so an HList sheet name has
'six characters less room than a VList one.
Friend Sub ReportSheetNameLength()
    Dim allowed As Long

    allowed = SHEET_NAME_MAX_LENGTH
    If this.layer = LLDataEntryLayerHList Then
        allowed = SHEET_NAME_MAX_LENGTH - Len(PRINT_PREFIX)
    End If

    If Len(this.sheetName) <= allowed Then Exit Sub

    LogInfo "The dictionary sheet name " & this.sheetName & " is " & _
            CStr(Len(this.sheetName)) & " characters long and this kind of data " & _
            "entry sheet allows " & CStr(allowed) & ", so its companion " & _
            "worksheets are created under shortened names. Shorten the sheet " & _
            "name in the dictionary to choose those names yourself.", _
            checkingWarning
End Sub


'@section Core
'===============================================================================
'@description Public entry point for building the data entry sheet.

'@label:build
'@sub-title Build the data entry sheet
'@details
'Dispatches to the private builder of the layer set during factory
'construction. HList mode delegates to BuildHList; VList mode delegates to
'BuildVList.
'@throws ProjectError.InvalidArgument When the layer value is not
'  LLDataEntryLayerHList or LLDataEntryLayerVList.
'@export
Public Sub Build()
    Select Case this.layer
        Case LLDataEntryLayerHList
            BuildHList
        Case LLDataEntryLayerVList
            BuildVList
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "There is no data entry layer '" & CStr(this.layer) & "' in a linelist"
    End Select
End Sub


'@section HListBuild
'===============================================================================
'@description Private implementation for horizontal data entry sheet construction.

'@label:build-hlist
'@sub-title Build a horizontal data entry sheet
'@details
'Constructs a full HList sheet: applies design formatting, freezes panes,
'adds action buttons (Show/Hide, Add Rows, Clear Filters, Open Print, and
'optionally Geo), builds sections via SectionBuilder, creates the GoTo
'section dropdown, stores sheet metadata as hidden names, creates the main
'ListObject table plus printed and filtered companion sheets, adds a pivot
'table, and protects the sheet. When a template workbook is in use, buttons
'are skipped and rows 3-4 are hidden instead.
'
'A failure anywhere in that run reaches the handler at the foot of this
'routine, which protects the sheet, puts screen updating back, files the
'failure for the generation report and re-raises what it caught.
Private Sub BuildHList()
    Dim sh As Worksheet
    Dim printSh As Worksheet
    Dim crfSh As Worksheet
    Dim filtSh As Worksheet
    Dim llshs As LLSheets
    Dim specs As LinelistSpecs
    Dim design As LLFormat
    Dim drop As DropdownLists
    Dim pass As Passwords
    Dim pivotObj As CustomPivotTable
    Dim shapeTrads As TranslationObject
    Dim messages As TranslationObject
    Dim rng As Range
    Dim but As Buttons
    Dim Lo As ListObject
    Dim tableName As String
    Dim shRowStart As Long
    Dim shRowEnd As Long
    Dim shColStart As Long
    Dim shColEnd As Long
    Dim shHn As HiddenNames
    Dim printHiddenNames As HiddenNames
    Dim crfHiddenNames As HiddenNames
    Dim filtSheetName As String
    Dim blankRowCells As Long
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    On Error GoTo BuildFailed

    Set llshs = this.sheetInfo
    Set sh = this.ll.Wksh(this.sheetName)
    Set specs = this.ll.LinelistData
    Set design = specs.DesignFormat
    Set drop = this.ll.Dropdown
    Set pass = specs.Password
    Set pivotObj = this.ll.Pivots

    'Add printed companion sheet
    this.ll.AddOutputSheet this.sheetName, xlSheetHidden, sheetScope:=2
    Set printSh = this.ll.PrintWksh(this.sheetName)

    'Add the CRF companion sheet. It holds one variable per row, laid out from
    'the `crf index` column of the dictionary, and SectionBuilder and VarWriter
    'have written to it all along -- the sheet itself was the one piece nothing
    'created, so BuildSections resolved it to Nothing and every CRF write was
    'skipped. It is created very hidden, which is the state ClickCloseSheet
    'leaves it in.
    this.ll.AddOutputSheet this.sheetName, xlSheetVeryHidden, sheetScope:=3
    Set crfSh = this.ll.CRFWksh(this.sheetName)

    'Format sheets
    design.ApplyFormat sh, scope:=AllLinelistSheet
    design.ApplyFormat printSh, scope:=HListPrintedSheet
    design.ApplyFormat crfSh, scope:=HListCRFSheet

    'Freeze panes
    FreezeHeader sh, llshs.DataBounds(this.sheetName, SheetBound.RowSart), 2

    'Add buttons (unless template workbook)
    If Not specs.HasTemplate() Then
        Set shapeTrads = specs.TransObject(TranslationOfShapes)

        'Show/Hide button
        Set rng = sh.Cells(1, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_SHOW_HIDE, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickShowHide", shapeLabel:=shapeTrads.TranslatedValue(SHP_SHOW_HIDE)
        but.Format design
        If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues

        'Add Rows button
        Set rng = sh.Cells(2, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_ADD_ROWS, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickAddRows", shapeLabel:=shapeTrads.TranslatedValue(SHP_ADD_ROWS)
        but.Format design
        If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues

        'Clear Filters button
        Set rng = sh.Cells(2, 1)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_REMOVE_FILT, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickRemoveFilters", shapeLabel:=shapeTrads.TranslatedValue(SHP_REMOVE_FILT)
        but.Format design
        If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues

        'Open Print button
        Set rng = sh.Cells(3, 1)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_OPEN_PRINT, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickOpenPrint", shapeLabel:=shapeTrads.TranslatedValue(SHP_OPEN_PRINT)
        but.Format design
        If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues

        'Geo button (if sheet has geo variables)
        If llshs.ContainsControl(this.sheetName, "geo1") Then
            Set rng = sh.Cells(3, 2)
            Set but = Buttons.Create(outrng:=rng, codeName:=SHP_GEO, scope:=ButtonScopeGeo)
            but.Add actionCommand:="ClickGeoApp", shapeLabel:=shapeTrads.TranslatedValue(SHP_GEO)
            but.Format design
            If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues
        End If

        'Show/Hide button on print sheet too
        Set rng = printSh.Cells(1, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_SHOW_HIDE, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickShowHide", shapeLabel:=shapeTrads.TranslatedValue(SHP_SHOW_HIDE)
        but.Format design
        If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues

        'And on the CRF sheet. The CRF is a show/hide layer of its own, with its
        'own `crf status` column, so without this button a workbook built with
        'no ribbon has no way to reach it.
        Set rng = crfSh.Cells(1, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_SHOW_HIDE, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickShowHide", shapeLabel:=shapeTrads.TranslatedValue(SHP_SHOW_HIDE)
        but.Format design
        If but.HasCheckings Then HarvestCollaboratorCheckings but.CheckingValues
    Else
        'Hide rows 3-4 when using ribbon template
        sh.Rows("3:4").EntireRow.Hidden = True
    End If

    'Build sections via SectionBuilder
    BuildSections sh, printSh, crfSh

    'GoTo section dropdown and sheet metadata
    Set messages = specs.TransObject(TranslationOfMessages)
    tableName = llshs.SheetInfo(this.sheetName, SheetInfoSheetTable)

    Set shHn = HiddenNames.Create(sh)
    BuildGoToSection sh, shHn, tableName, design, drop, messages
    WriteSheetMetadata shHn, "HList", tableName

    Set printHiddenNames = HiddenNames.Create(printSh)
    WriteSheetMetadata printHiddenNames, "HList Print", PRINT_PREFIX & tableName

    'The CRF tag is what ClickOpenCRF, ClickCloseSheet and ResolveShowHideLayer
    'all read to tell the sheet apart from its two siblings.
    Set crfHiddenNames = HiddenNames.Create(crfSh)
    WriteSheetMetadata crfHiddenNames, "HList CRF", CRF_PREFIX & tableName

    'Create the main ListObject table
    With sh
        shRowStart = llshs.DataBounds(this.sheetName, SheetBound.RowSart)
        shRowEnd = shRowStart + 1
        shColStart = llshs.DataBounds(this.sheetName, SheetBound.ColStart)
        shColEnd = llshs.DataBounds(this.sheetName, SheetBound.ColEnd)

        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = tableName

        'Resize to full data bounds
        Set Lo = .ListObjects(tableName)
        shRowEnd = llshs.DataBounds(this.sheetName, SheetBound.RowEnd)
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        'The calculation before the resize has never been measured. It costs a
        'full recalculation of this worksheet per data entry sheet built, and
        'whether Resize needs it is open.
        .Calculate
        Lo.Resize rng

        design.ApplyFormat Lo, scope:=HListTable
        Set rng = Lo.HeaderRowRange
        design.ApplyFormat rng, scope:=HListHeader
    End With

    'How many cells an untouched data row carries. Every reader compares a row
    'against this to decide whether the row holds anything a user typed.
    With Lo.ListRows(1)
        blankRowCells = Application.WorksheetFunction.CountA(.Range)
        EnsureAndSet shHn, "blank_row_count", blankRowCells, HiddenNameTypeLong
        EnsureAndSet printHiddenNames, "blank_row_count", blankRowCells, HiddenNameTypeLong
    End With

    'Add pivot table for this sheet. Add is what marks the pivot sheet as
    'carrying content and Format is what reads that mark, so the order holds.
    pivotObj.Add this.sheetName, tableName, messages.TranslatedValue("MSG_PivotTable")
    pivotObj.Format design
    If pivotObj.HasCheckings Then HarvestCollaboratorCheckings pivotObj.CheckingValues

    'Printed companion ListObject
    With printSh
        shRowEnd = shRowStart + 1
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = PRINT_PREFIX & tableName

        Set Lo = .ListObjects(PRINT_PREFIX & tableName)
        shRowEnd = shRowEnd + COMPANION_SPARE_ROWS
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        Lo.Resize rng
        design.ApplyFormat Lo, scope:=HListTable

        Set rng = Lo.HeaderRowRange
        design.ApplyFormat rng, scope:=HListPrintedHeader
    End With

    'Filtered companion sheet. It takes the header row and room to grow, the way
    'the printed companion above does. ExportButton resizes it to the address of
    'the source table each time it refreshes the filtered copy.
    filtSheetName = FILTERED_PREFIX & this.sheetName
    this.ll.AddOutputSheet filtSheetName, xlSheetVeryHidden
    Set filtSh = this.ll.Wksh(filtSheetName)

    Set Lo = sh.ListObjects(tableName)
    With filtSh
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowStart, shColEnd))
        rng.Value = Lo.HeaderRowRange.Value

        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowStart + 1, shColEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = FILTERED_PREFIX & tableName

        Set Lo = .ListObjects(FILTERED_PREFIX & tableName)
        Set rng = .Range(.Cells(shRowStart, shColStart), _
                         .Cells(shRowStart + 1 + COMPANION_SPARE_ROWS, shColEnd))
        Lo.Resize rng
    End With

    EnsureAndSet shHn, "filtered_sheet", filtSh.Name

    pass.Protect sh

    LogInfo "HList sheet " & this.sheetName & " built (table: " & tableName & _
            ", " & BuiltCountsText() & ")"
    Exit Sub

BuildFailed:
    'Copy the error before anything else runs. The cleanup below arms its own
    'On Error, and that clears Err.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description
    Resume CleanUpAfterFail

CleanUpAfterFail:
    On Error Resume Next
        If Not (sh Is Nothing Or pass Is Nothing) Then pass.Protect sh
        LogInfo "The HList sheet " & this.sheetName & " could not be built: " & _
                failedDescription, checkingError
    On Error GoTo 0

    Err.Raise failedNumber, failedSource, failedDescription
End Sub


'@section VListBuild
'===============================================================================
'@description Private implementation for vertical data entry sheet construction.

'@label:build-vlist
'@sub-title Build a vertical data entry sheet
'@details
'Constructs a VList sheet: applies design formatting, freezes pane row 1,
'builds sections via SectionBuilder, creates the GoTo section dropdown,
'stores sheet metadata as hidden names, computes the PLAGEVALUES range
'reference and protects the sheet.
'
'A failure anywhere in that run reaches the handler at the foot of this
'routine, which protects the sheet, puts screen updating back, files the
'failure for the generation report and re-raises what it caught.
Private Sub BuildVList()
    Dim sh As Worksheet
    Dim specs As LinelistSpecs
    Dim design As LLFormat
    Dim drop As DropdownLists
    Dim pass As Passwords
    Dim messages As TranslationObject
    Dim tableName As String
    Dim shHn As HiddenNames
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    On Error GoTo BuildFailed

    Set sh = this.ll.Wksh(this.sheetName)
    Set specs = this.ll.LinelistData
    Set design = specs.DesignFormat
    Set drop = this.ll.Dropdown
    Set pass = specs.Password

    'Format the sheet
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Freeze panes (row 1 only)
    FreezeHeader sh, 1

    'Build sections via SectionBuilder
    BuildSections sh, Nothing

    'GoTo section dropdown and sheet metadata
    Set messages = specs.TransObject(TranslationOfMessages)
    tableName = this.sheetInfo.SheetInfo(this.sheetName, SheetInfoSheetTable)

    Set shHn = HiddenNames.Create(sh)
    BuildGoToSection sh, shHn, tableName, design, drop, messages
    WriteSheetMetadata shHn, "VList", tableName

    NameValuesRange sh, tableName

    pass.Protect sh

    LogInfo "VList sheet " & this.sheetName & " built (table: " & tableName & _
            ", " & BuiltCountsText() & ")"
    Exit Sub

BuildFailed:
    'Copy the error before anything else runs. The cleanup below arms its own
    'On Error, and that clears Err.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description
    Resume CleanUpAfterFail

CleanUpAfterFail:
    On Error Resume Next
        If Not (sh Is Nothing Or pass Is Nothing) Then pass.Protect sh
        LogInfo "The VList sheet " & this.sheetName & " could not be built: " & _
                failedDescription, checkingError
    On Error GoTo 0

    Err.Raise failedNumber, failedSource, failedDescription
End Sub


'@section Helpers
'===============================================================================
'@description Private utilities shared by both build modes.

'@label:freeze-header
'@sub-title Freeze the header of a data entry sheet
'@details
'ActiveWindow belongs to whichever sheet is active, so the activation is tested
'before the three writes. A sheet that refuses to activate - a hidden one, or a
'workbook whose structure is protected - used to have the freeze applied to
'whatever worksheet was active at the time, with no message.
'@param sh Worksheet. The sheet whose panes are frozen.
'@param splitRow Long. The row the horizontal split sits under.
'@param splitColumn Optional Long. The column the vertical split sits after.
'  Zero leaves the columns unsplit.
Private Sub FreezeHeader(ByVal sh As Worksheet, ByVal splitRow As Long, _
                         Optional ByVal splitColumn As Long = 0)
    On Error Resume Next
        sh.Activate
    On Error GoTo 0

    If Not (ActiveSheet Is sh) Then
        LogInfo "The panes of " & this.sheetName & " were left unfrozen: the " & _
                "worksheet could not be brought to the front.", checkingWarning
        Exit Sub
    End If

    If ActiveWindow Is Nothing Then
        LogInfo "The panes of " & this.sheetName & " were left unfrozen: the " & _
                "workbook has no window.", checkingWarning
        Exit Sub
    End If

    With ActiveWindow
        .SplitRow = splitRow
        If splitColumn > 0 Then .SplitColumn = splitColumn
        .FreezePanes = True
    End With
End Sub

'@label:build-sections
'@sub-title Build sections using SectionBuilder
'@details
'Creates a SectionBuilder in the appropriate mode (HList or VList) and
'invokes its Build method. Both standard and custom dropdown objects are
'passed to the builder for variable validation setup.
'
'The CRF companion arrives as a parameter. It used to be looked up here under
'a suppressed error, which read as "use it when it happens to be there" and
'hid the fact that nothing ever created it.
'@param sh Worksheet. The main data entry sheet to build sections on.
'@param printSh Worksheet. The printed companion sheet (Nothing for VList).
'@param crfSh Worksheet. The CRF companion sheet (Nothing for VList).
Private Sub BuildSections(ByVal sh As Worksheet, _
                          ByVal printSh As Worksheet, _
                          Optional ByVal crfSh As Worksheet = Nothing)
    Dim specs As LinelistSpecs
    Dim secBuilder As SectionBuilder
    Dim dropdownObj As DropdownLists
    Dim customDropdownObj As DropdownLists

    Set specs = this.ll.LinelistData
    Set dropdownObj = this.ll.Dropdown(1)
    Set customDropdownObj = this.ll.Dropdown(2)

    If this.layer = LLDataEntryLayerHList Then
        Set secBuilder = SectionBuilder.Create( _
            SectionBuilderModeHList, specs, sh, printSh, crfSh, dropdownObj, customDropdownObj)
    Else
        Set secBuilder = SectionBuilder.Create( _
            SectionBuilderModeVList, specs, sh, , , dropdownObj, customDropdownObj)
    End If

    secBuilder.Build this.sheetName, SectionStartRow()

    'One section builder serves the whole sheet, so the field is what lets
    'the caller read the counts and the per-variable record after the build.
    Set this.builder = secBuilder

    'Harvest per-variable checkings from the section builder
    If secBuilder.HasCheckings Then HarvestCollaboratorCheckings secBuilder.CheckingValues
End Sub

'@label:build-goto-section
'@sub-title Write the GoTo section dropdown into cell A1
'@details
'The same twenty lines used to sit in each builder, and the two copies of the
'dictionary index they carried drifted apart by one row. One copy is what keeps
'them together.
'@param sh Worksheet. The data entry sheet.
'@param shHn HiddenNames. The hidden name store of that sheet.
'@param tableName String. The sheet's table name from the dictionary.
'@param design LLFormat. The design format of the linelist.
'@param drop DropdownLists. The standard dropdown manager of the linelist.
'@param messages TranslationObject. The message translation scope.
Private Sub BuildGoToSection(ByVal sh As Worksheet, _
                             ByVal shHn As HiddenNames, _
                             ByVal tableName As String, _
                             ByVal design As LLFormat, _
                             ByVal drop As DropdownLists, _
                             ByVal messages As TranslationObject)
    Dim rng As Range
    Dim sectionsList As BetterArray
    Dim goToSection As String
    Dim gotoSectionListName As String

    goToSection = messages.TranslatedValue("MSG_GoToSec")
    Set sectionsList = CollectSectionNames(goToSection)
    gotoSectionListName = tableName & "_" & GOTO_SECTION_CODE

    Set rng = sh.Cells(1, 1)
    EnsureAndSet shHn, gotoSectionListName, goToSection
    rng.Value = goToSection

    drop.Add sectionsList, gotoSectionListName
    drop.SetValidation rng, gotoSectionListName
    design.ApplyFormat rng, scope:=LinelistSelectionDropdown
    rng.Locked = False
End Sub

'@label:write-sheet-metadata
'@sub-title Record the kind of sheet this is and the table it carries
'@param store HiddenNames. The hidden name store of the sheet.
'@param sheetType String. "HList", "HList Print" or "VList".
'@param tableName String. The name of the sheet's table.
Private Sub WriteSheetMetadata(ByVal store As HiddenNames, _
                               ByVal sheetType As String, _
                               ByVal tableName As String)
    EnsureAndSet store, "sheet_type", sheetType
    EnsureAndSet store, "table_name", tableName
End Sub

'@label:name-values-range
'@sub-title Name the block of value cells of a VList sheet
'@details
'The walk starts one column left of the table marker, which is the column
'carrying the variable labels, and runs down until it meets an empty cell. The
'named range ends on that empty cell, so it holds one row past the last value:
'that row is where the next value lands, and EventLinelist recalculates the
'whole named range when a cell in it changes.
'
'The walk is bounded by the last row of the worksheet. A marker sitting in
'column A is reported, because the column to its left is off the sheet.
'@param sh Worksheet. The VList data entry sheet.
'@param tableName String. The sheet's table name from the dictionary.
Private Sub NameValuesRange(ByVal sh As Worksheet, ByVal tableName As String)
    Dim cellRng As Range
    Dim rng As Range
    Dim lastRow As Long
    Dim walkRow As Long

    Set cellRng = sh.Range(tableName & "_START")

    If cellRng.Column <= 1 Then
        LogInfo "The values of " & this.sheetName & " have no named range: the " & _
                "table marker sits in column A and the labels of a vertical " & _
                "sheet are read from the column to its left.", checkingError
        Exit Sub
    End If

    lastRow = sh.Rows.Count
    walkRow = cellRng.Row

    Do While walkRow < lastRow
        If IsEmpty(sh.Cells(walkRow, cellRng.Column - 1).Value) Then Exit Do
        walkRow = walkRow + 1
    Loop

    Set rng = sh.Range(sh.Cells(cellRng.Row, cellRng.Column), _
                       sh.Cells(walkRow, cellRng.Column))
    rng.Name = tableName & "_PLAGEVALUES"
End Sub

'@label:section-start-row
'@sub-title The dictionary index of this sheet's first variable
'@details
'LLSheets.RowIndex answers an absolute worksheet row. The dictionary data range
'is taken without its header and the linelist dictionary keeps that header on
'worksheet row 1, so worksheet row R is index R - 1. Both readers of the
'dictionary in this class go through here, which is what holds them together:
'the two used to differ by one and a sheet whose first section held a single
'variable lost that section from its GoTo dropdown.
'@return Long. The 1-based index into the dictionary data range.
Private Function SectionStartRow() As Long
    SectionStartRow = this.sheetInfo.RowIndex(this.sheetName) - 1
End Function

'@label:collect-section-names
'@sub-title Collect section names for the GoTo dropdown
'@details
'Walks the dictionary rows of the current sheet and keeps each new section name
'in the order it appears, prefixed with the GoTo label. A section that repeats
'on consecutive rows is kept once.
'
'The two columns are read into memory in one crossing each. Reading them a cell
'at a time cost two crossings per dictionary row, per sheet built.
'@param goToLabel String. The translated "Go to section" prefix label.
'@return BetterArray. A 1-based array of prefixed section name strings.
Private Function CollectSectionNames(ByVal goToLabel As String) As BetterArray
    Dim dict As LLdictionary
    Dim sectionValues As Variant
    Dim sheetValues As Variant
    Dim startRow As Long
    Dim lastRow As Long
    Dim counter As Long
    Dim sectionName As String
    Dim prevSection As String
    Dim result As BetterArray

    Set result = New BetterArray
    result.LowerBound = 1

    Set dict = this.ll.Dictionary
    sectionValues = ColumnOf(dict, "main section")
    sheetValues = ColumnOf(dict, "sheet name")

    startRow = SectionStartRow()
    If startRow < 1 Then startRow = 1

    lastRow = startRow + this.sheetInfo.NumberOfVars(this.sheetName) - 1
    If lastRow > UBound(sheetValues) Then lastRow = UBound(sheetValues)
    If lastRow > UBound(sectionValues) Then lastRow = UBound(sectionValues)

    prevSection = vbNullString

    For counter = startRow To lastRow
        If CStr(sheetValues(counter)) = this.sheetName Then
            sectionName = CStr(sectionValues(counter))
            If sectionName <> vbNullString And sectionName <> prevSection Then
                result.Push goToLabel & ": " & sectionName
                prevSection = sectionName
            End If
        End If
    Next

    Set CollectSectionNames = result
End Function

'@label:column-of
'@sub-title Read one dictionary column into a 1-based array
'@details
'A one-cell range answers a scalar, and a column the dictionary does not carry
'answers Nothing. Both come back as a one-entry array holding an empty string,
'so the caller walks the same shape in every case.
'@param dict LLdictionary. The linelist dictionary.
'@param columnName String. The dictionary column to read.
'@return Variant. A 1-based array of the column's values.
Private Function ColumnOf(ByVal dict As LLdictionary, ByVal columnName As String) As Variant
    Dim rng As Range
    Dim values As Variant
    Dim flat() As Variant
    Dim counter As Long

    Set rng = dict.DataRange(columnName)

    If rng Is Nothing Then
        ReDim flat(1 To 1)
        flat(1) = vbNullString
        ColumnOf = flat
        Exit Function
    End If

    If rng.Cells.Count = 1 Then
        ReDim flat(1 To 1)
        flat(1) = rng.Value
        ColumnOf = flat
        Exit Function
    End If

    values = rng.Value
    ReDim flat(1 To UBound(values, 1))
    For counter = 1 To UBound(values, 1)
        flat(counter) = values(counter, 1)
    Next

    ColumnOf = flat
End Function

'@label:ensure-and-set
'@sub-title Give a hidden name its value, creating it when it is absent
'@details
'HiddenNames.EnsureName writes the value only when it creates the name, so a
'name that is already there keeps whatever it held. AddOutputSheet reuses a
'worksheet a template workbook already carries, and every one of these names is
'read by other classes, so a second build has to overwrite what the first one
'wrote.
'@param store HiddenNames. The hidden name store.
'@param nameId String. The hidden name identifier.
'@param value Variant. The value to write.
'@param valueType Optional Byte. The stored type. Defaults to a string.
Private Sub EnsureAndSet(ByVal store As HiddenNames, ByVal nameId As String, _
                         ByVal value As Variant, _
                         Optional ByVal valueType As Byte = HiddenNameTypeString)
    If store.HasName(nameId) Then
        store.SetValue nameId, value
    Else
        store.EnsureName nameId, value, valueType
    End If
End Sub


'@section Checkings
'===============================================================================

'@label:log-info
'@sub-title File one entry against the sheet being built
'@details
'The key carries the sheet name, so the entries of two sheets can be merged
'into one Checking. The counter alone started at 0 in every instance and
'`Checking.Add` raises on a duplicate, so the second sheet of a run stopped the
'caller. `SectionBuilder` keys its entries by worksheet and variable for the
'same reason.
'@param label String. The message for the report.
'@param scope Optional Byte. Severity. Defaults to checkingSuccess.
Private Sub LogInfo(ByVal label As String, Optional ByVal scope As Byte = checkingSuccess)
    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create("Sheet build: " & this.sheetName)
    End If
    this.internalChecks.Add this.sheetName & " #" & CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:harvest-collaborator-checkings
'@sub-title Fold what a collaborator filed into this sheet's entries
'@details
'The two shared dropdown managers are left out on purpose. Linelist keeps one
'instance of each for the whole build and the generation report flushes them
'once, after every sheet is built, so folding them in here would report each of
'their warnings once per data entry sheet as well.
'@param collaboratorChecks Checking. What the collaborator filed.
Private Sub HarvestCollaboratorCheckings(ByVal collaboratorChecks As Checking)
    If collaboratorChecks Is Nothing Then Exit Sub
    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create("Sheet build: " & this.sheetName)
    End If
    this.internalChecks.Append collaboratorChecks
    this.hasChecks = True
End Sub

'@label:built-counts-text
'@sub-title The section and variable counts of the finished sheet
'@details
'Read off the section builder the sheet was laid out with. A sheet whose
'sections were never built answers zeros, which is what the counts are.
'@return String. The counts as one phrase for the report line.
Private Function BuiltCountsText() As String
    BuiltCountsText = CStr(SectionsPlaced()) & " section(s), " & _
                      CStr(VariablesWritten()) & " variable(s)"
End Function

'@label:has-milestones
'@prop-title Whether any variable of this sheet was written
'@return Boolean. True once a variable of this sheet reached the worksheet.
'@export
Public Property Get HasMilestones() As Boolean
    If this.builder Is Nothing Then Exit Property
    HasMilestones = this.builder.HasMilestones
End Property

'@label:milestone-values
'@prop-title The per-variable record of this sheet
'@details
'Handed up from the section builder as the writer holds it. The build
'driver takes it into the run log record-only, so it reaches the text
'file and leaves the __check worksheet to the sheet-level entries and
'the problems.
'@return Checking. The milestone entries, or Nothing.
'@export
Public Property Get MilestoneValues() As Checking
    If this.builder Is Nothing Then Exit Property
    Set MilestoneValues = this.builder.MilestoneValues
End Property

'@label:variables-written
'@prop-title How many variables this sheet build put on the worksheet
'@return Long. The number of variables written.
'@export
Public Property Get VariablesWritten() As Long
    If this.builder Is Nothing Then Exit Property
    VariablesWritten = this.builder.VariablesWritten
End Property

'@label:sections-placed
'@prop-title How many sections this sheet build laid out
'@return Long. The number of sections placed.
'@export
Public Property Get SectionsPlaced() As Long
    If this.builder Is Nothing Then Exit Property
    SectionsPlaced = this.builder.SectionsPlaced
End Property

'@label:has-checkings
'@prop-title Whether this build has anything to report
'@details
'Answers True once an entry has been filed. The caller tests this before
'reading CheckingValues and pushes the result into the generation report.
'@return Boolean. True when there is something to read.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title The report entries of this sheet build
'@details
'Answers the Checking object holding the entries, and Nothing while none have
'been filed. EventsDesignerAdvanced flushes it as phase 2 of the generation
'report, once every data entry sheet is built.
'@return Checking. The entries, or Nothing.
'@export
Public Property Get CheckingValues() As Checking
    If Not HasCheckings Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@section Errors
'===============================================================================

'@label:guard-not-sealed
'@sub-title Refuse a write to a creation-only binding after Create
'@param propName String. The binding the caller tried to write.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. The error code to raise.
'@param errorMessage String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, errorMessage
End Sub
