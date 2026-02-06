VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Builds HList or VList data entry sheets with sections, formatting, and metadata"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Builds HList or VList data entry sheets with sections, formatting, and metadata")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

Implements IListBuilder

Private Const CLASS_NAME As String = "ListBuilder"
Private Const GOTO_SECTION_CODE As String = "go_to_section"
Private Const FILTERED_PREFIX As String = "f"
Private Const PRINT_PREFIX As String = "print_"

'Button shape codes
Private Const SHP_GEO As String = "SHP_GeoApps"
Private Const SHP_ADD_ROWS As String = "SHP_Add200L"
Private Const SHP_REMOVE_FILT As String = "SHP_ClearFilters"
Private Const SHP_SHOW_HIDE As String = "SHP_ShowHide"
Private Const SHP_OPEN_PRINT As String = "SHP_OpenPrint"

Private Type TListBuilder
    layer As Byte
    sheetName As String
    ll As ILinelist
End Type

Private this As TListBuilder

'@section Factory
'===============================================================================

Public Function Create(ByVal layer As Byte, _
                       ByVal sheetName As String, _
                       ByVal ll As ILinelist) As IListBuilder

    If ll Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist object is required"
    End If

    If LenB(Trim$(sheetName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheet name cannot be empty"
    End If

    Dim llshs As ILLSheets
    Set llshs = LLSheets.Create(ll.Dictionary)

    If Not llshs.Contains(sheetName) Then
        ThrowError ProjectError.InvalidArgument, _
                   "The sheet name " & sheetName & " is not present in the dictionary"
    End If

    Dim instance As ListBuilder
    Set instance = New ListBuilder

    instance.BinLayer = layer
    instance.BinSheetName = sheetName
    Set instance.BinLinelist = ll

    Set Create = instance
End Function

'@section Bindings
'===============================================================================

Friend Property Let BinLayer(ByVal value As Byte)
    this.layer = value
End Property

Friend Property Let BinSheetName(ByVal value As String)
    this.sheetName = value
End Property

Friend Property Set BinLinelist(ByVal value As ILinelist)
    Set this.ll = value
End Property

'@section Core
'===============================================================================

Public Sub Build()
    Select Case this.layer
        Case ListBuilderLayerHList
            BuildHList
        Case ListBuilderLayerVList
            BuildVList
        Case Else
            ThrowError ProjectError.InvalidArgument, _
                       "Unknown list builder layer: " & CStr(this.layer)
    End Select
End Sub

'@section Private — HList build
'===============================================================================

Private Sub BuildHList()
    Dim sh As Worksheet
    Dim printSh As Worksheet
    Dim filtSh As Worksheet
    Dim dict As ILLdictionary
    Dim llshs As ILLSheets
    Dim specs As ILinelistSpecs
    Dim design As ILLFormat
    Dim drop As IDropdownLists
    Dim pass As IPasswords
    Dim mainObj As IMain
    Dim pivotObj As ICustomPivotTables
    Dim trads As ITranslationObject
    Dim rng As Range
    Dim but As IButtons
    Dim Lo As ListObject
    Dim sectionsList As BetterArray
    Dim tableName As String
    Dim gotoSectionListName As String
    Dim goToSection As String
    Dim shRowStart As Long
    Dim shRowEnd As Long
    Dim shColStart As Long
    Dim shColEnd As Long
    Dim hasGeo As Boolean
    Dim hiddenNames As IHiddenNames
    Dim filtSheetName As String
    Dim countaValue As Long

    Set sh = this.ll.Wksh(this.sheetName)
    Set dict = this.ll.Dictionary
    Set llshs = LLSheets.Create(dict)
    Set specs = this.ll.LinelistData
    Set design = specs.DesignFormat
    Set drop = this.ll.Dropdown
    Set pass = specs.Password
    Set mainObj = specs.MainObject
    Set pivotObj = this.ll.Pivots

    'Add printed companion sheet
    this.ll.AddOutputSheet this.sheetName, xlSheetHidden, sheetScope:=2
    Set printSh = this.ll.PrintWksh(this.sheetName)

    'Format sheets
    design.ApplyFormat sh, scope:=AllLinelistSheet
    design.ApplyFormat printSh, scope:=HListPrintedSheet

    'Freeze panes
    On Error Resume Next
    sh.Activate
    BusyApp
    With ActiveWindow
        .SplitRow = llshs.DataBounds(this.sheetName, 1)
        .SplitColumn = 2
        .FreezePanes = True
    End With
    On Error GoTo 0

    'Add buttons (unless template workbook)
    If Not mainObj.HasTemplate() Then
        Set trads = specs.TransObject(TranslationOfShapes)

        'Show/Hide button
        Set rng = sh.Cells(1, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_SHOW_HIDE, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickShowHide", shapeLabel:=trads.TranslatedValue(SHP_SHOW_HIDE)
        but.Format design

        'Add Rows button
        Set rng = sh.Cells(2, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_ADD_ROWS, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickAddRows", shapeLabel:=trads.TranslatedValue(SHP_ADD_ROWS)
        but.Format design

        'Clear Filters button
        Set rng = sh.Cells(2, 1)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_REMOVE_FILT, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickRemoveFilters", shapeLabel:=trads.TranslatedValue(SHP_REMOVE_FILT)
        but.Format design

        'Open Print button
        Set rng = sh.Cells(3, 1)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_OPEN_PRINT, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickOpenPrint", shapeLabel:=trads.TranslatedValue(SHP_OPEN_PRINT)
        but.Format design

        'Geo button (if sheet has geo variables)
        hasGeo = llshs.ContainsControl(this.sheetName, "geo1")
        If hasGeo Then
            Set rng = sh.Cells(3, 2)
            Set but = Buttons.Create(outrng:=rng, codeName:=SHP_GEO, scope:=ButtonScopeGeo)
            but.Add actionCommand:="ClickGeoApp", shapeLabel:=trads.TranslatedValue(SHP_GEO)
            but.Format design
        End If

        'Show/Hide button on print sheet too
        Set rng = printSh.Cells(1, 2)
        Set but = Buttons.Create(outrng:=rng, codeName:=SHP_SHOW_HIDE, scope:=ButtonScopeSmall)
        but.Add actionCommand:="ClickShowHide", shapeLabel:=trads.TranslatedValue(SHP_SHOW_HIDE)
        but.Format design
    Else
        'Hide rows 3-4 when using ribbon template
        sh.Rows("3:4").EntireRow.Hidden = True
    End If

    'Build sections via SectionBuilder
    BuildSections sh, printSh

    'GoTo section dropdown
    Set trads = specs.TransObject(TranslationOfMessages)
    goToSection = trads.TranslatedValue("MSG_GoToSec")
    Set sectionsList = CollectSectionNames(dict, llshs, goToSection)

    tableName = llshs.SheetInfo(this.sheetName, SheetInfoSheetTable)
    gotoSectionListName = tableName & "_" & GOTO_SECTION_CODE

    Set rng = sh.Cells(1, 1)
    Set hiddenNames = HiddenNames.Create(sh)
    hiddenNames.EnsureName gotoSectionListName, goToSection, HiddenNameTypeString
    rng.Value = goToSection

    drop.Add sectionsList, gotoSectionListName
    drop.SetValidation rng, gotoSectionListName
    design.ApplyFormat rng, scope:=LinelistSelectionDropdown
    rng.Locked = False

    'Store sheet metadata via hidden names
    hiddenNames.EnsureName "sheet_type", "HList", HiddenNameTypeString
    hiddenNames.EnsureName "table_name", tableName, HiddenNameTypeString

    'Print sheet hidden names
    Dim printHiddenNames As IHiddenNames
    Set printHiddenNames = HiddenNames.Create(printSh)
    printHiddenNames.EnsureName "sheet_type", "HList Print", HiddenNameTypeString
    printHiddenNames.EnsureName "table_name", PRINT_PREFIX & tableName, HiddenNameTypeString

    'Create the main ListObject table
    With sh
        shRowStart = llshs.DataBounds(this.sheetName, 1)
        shRowEnd = shRowStart + 1
        shColStart = llshs.DataBounds(this.sheetName, 3)
        shColEnd = llshs.DataBounds(this.sheetName, 4)

        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = tableName

        'Resize to full data bounds
        Set Lo = .ListObjects(tableName)
        shRowEnd = llshs.DataBounds(this.sheetName, 2)
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        .Calculate
        Lo.Resize rng

        design.ApplyFormat Lo, scope:=HListTable
        Set rng = Lo.HeaderRowRange
        design.ApplyFormat rng, scope:=HListHeader
    End With

    'Store row count via hidden names
    With Lo.ListRows(1)
        countaValue = Application.WorksheetFunction.CountA(.Range)
        hiddenNames.EnsureName "row_count", CStr(countaValue), HiddenNameTypeLong
        printHiddenNames.EnsureName "row_count", CStr(countaValue), HiddenNameTypeLong
    End With

    'Add pivot table for this sheet
    pivotObj.Add this.sheetName, tableName, trads.TranslatedValue("MSG_PivotTable")
    pivotObj.Format design

    'Printed companion ListObject
    With printSh
        shRowEnd = shRowStart + 1
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = PRINT_PREFIX & tableName

        Set Lo = .ListObjects(PRINT_PREFIX & tableName)
        shRowEnd = shRowEnd + 10
        Set rng = .Range(.Cells(shRowStart, shColStart), .Cells(shRowEnd, shColEnd))
        Lo.Resize rng
        design.ApplyFormat Lo, scope:=HListTable

        Set rng = Lo.HeaderRowRange
        design.ApplyFormat rng, scope:=HListPrintedHeader
    End With

    'Filtered companion sheet
    filtSheetName = FILTERED_PREFIX & this.sheetName
    this.ll.AddOutputSheet filtSheetName, xlSheetVeryHidden
    Set filtSh = this.ll.Wksh(filtSheetName)

    Set Lo = sh.ListObjects(tableName)
    With filtSh
        Set rng = .Range(Lo.Range.Address)
        rng.Value = Lo.Range.Value
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = FILTERED_PREFIX & tableName
    End With

    hiddenNames.EnsureName "filtered_sheet", filtSh.Name, HiddenNameTypeString

    pass.Protect sh
End Sub

'@section Private — VList build
'===============================================================================

Private Sub BuildVList()
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim llshs As ILLSheets
    Dim specs As ILinelistSpecs
    Dim design As ILLFormat
    Dim drop As IDropdownLists
    Dim pass As IPasswords
    Dim trads As ITranslationObject
    Dim rng As Range
    Dim cellRng As Range
    Dim sectionsList As BetterArray
    Dim tableName As String
    Dim gotoSectionListName As String
    Dim goToSection As String
    Dim hiddenNames As IHiddenNames

    Set sh = this.ll.Wksh(this.sheetName)
    Set dict = this.ll.Dictionary
    Set llshs = LLSheets.Create(dict)
    Set specs = this.ll.LinelistData
    Set design = specs.DesignFormat
    Set drop = this.ll.Dropdown
    Set pass = specs.Password

    'Format the sheet
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Freeze panes (row 1 only)
    On Error Resume Next
    sh.Activate
    BusyApp
    With ActiveWindow
        .SplitRow = 1
        .FreezePanes = True
    End With
    On Error GoTo 0

    'Build sections via SectionBuilder
    BuildSections sh, Nothing

    'GoTo section dropdown
    Set trads = specs.TransObject(TranslationOfMessages)
    goToSection = trads.TranslatedValue("MSG_GoToSec")
    Set sectionsList = CollectSectionNames(dict, llshs, goToSection)

    tableName = llshs.SheetInfo(this.sheetName, SheetInfoSheetTable)
    gotoSectionListName = tableName & "_" & GOTO_SECTION_CODE

    Set rng = sh.Cells(1, 1)
    Set hiddenNames = HiddenNames.Create(sh)
    hiddenNames.EnsureName gotoSectionListName, goToSection, HiddenNameTypeString
    rng.Value = goToSection

    drop.Add sectionsList, gotoSectionListName
    drop.SetValidation rng, gotoSectionListName
    design.ApplyFormat rng, scope:=LinelistSelectionDropdown
    rng.Locked = False

    'Store sheet metadata via hidden names
    hiddenNames.EnsureName "sheet_type", "VList", HiddenNameTypeString
    hiddenNames.EnsureName "table_name", tableName, HiddenNameTypeString

    'PLAGEVALUES range reference
    Set cellRng = sh.Range(tableName & "_START")
    Set rng = cellRng.Offset(, -1)

    Do While Not IsEmpty(rng)
        Set rng = rng.Offset(1)
    Loop

    With sh
        Set rng = .Range(.Cells(cellRng.Row, cellRng.Column), .Cells(rng.Row, cellRng.Column))
    End With

    hiddenNames.EnsureName tableName & "_PLAGEVALUES", rng.Address, HiddenNameTypeString

    pass.Protect "_active"
End Sub

'@section Private — Shared helpers
'===============================================================================

'Build sections using SectionBuilder (Phase 1 class).
Private Sub BuildSections(ByVal sh As Worksheet, ByVal printSh As Worksheet)
    Dim specs As ILinelistSpecs
    Dim llshs As ILLSheets
    Dim startRow As Long
    Dim secBuilder As ISectionBuilder
    Dim crfSh As Worksheet
    Dim dropdownObj As IDropdownLists
    Dim customDropdownObj As IDropdownLists

    Set specs = this.ll.LinelistData
    Set llshs = LLSheets.Create(this.ll.Dictionary)
    startRow = llshs.RowIndex(this.sheetName) - 1

    Set dropdownObj = this.ll.Dropdown(1)
    Set customDropdownObj = this.ll.Dropdown(2)

    If this.layer = ListBuilderLayerHList Then
        On Error Resume Next
        Set crfSh = this.ll.CRFWksh(this.sheetName)
        On Error GoTo 0

        Set secBuilder = SectionBuilder.Create( _
            SectionBuilderModeHList, specs, sh, printSh, crfSh, dropdownObj, customDropdownObj)
    Else
        Set secBuilder = SectionBuilder.Create( _
            SectionBuilderModeVList, specs, sh, , , dropdownObj, customDropdownObj)
    End If

    secBuilder.Build this.sheetName, startRow
End Sub

'Collect section names from the dictionary for the GoTo dropdown.
Private Function CollectSectionNames(ByVal dict As ILLdictionary, _
                                     ByVal llshs As ILLSheets, _
                                     ByVal goToLabel As String) As BetterArray
    Dim sectionRng As Range
    Dim sheetRng As Range
    Dim startRow As Long
    Dim numVars As Long
    Dim counter As Long
    Dim sectionName As String
    Dim prevSection As String
    Dim result As BetterArray

    Set result = New BetterArray
    result.LowerBound = 1

    Set sectionRng = dict.DataRange("section")
    Set sheetRng = dict.DataRange("sheet name")
    startRow = llshs.RowIndex(this.sheetName)
    numVars = llshs.NumberOfVars(this.sheetName)

    prevSection = vbNullString

    For counter = startRow To startRow + numVars - 1
        If CStr(sheetRng.Cells(counter, 1).Value) = this.sheetName Then
            sectionName = CStr(sectionRng.Cells(counter, 1).Value)
            If sectionName <> vbNullString And sectionName <> prevSection Then
                result.Push goToLabel & ": " & sectionName
                prevSection = sectionName
            End If
        End If
    Next

    Set CollectSectionNames = result
End Function

'Suppress UI updates for batch operations.
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Raise a project error with this class as the source.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, errorMessage
End Sub

'@section IListBuilder implementation
'===============================================================================

Private Sub IListBuilder_Build()
    Build
End Sub
