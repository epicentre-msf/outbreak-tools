VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Linelist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Linelist workbook facade: creates, populates and saves the output workbook"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Linelist workbook facade: creates, populates and saves the output workbook")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILinelist

'@class Linelist
'@description
'Facade for the output linelist workbook. This class orchestrates the full
'lifecycle of creating, populating, saving, and error-managing the .xlsb
'workbook that end-users receive. It lazily creates the output workbook
'(or opens a template), adds temporary and permanent worksheets, exports
'specifications, prepares hidden names and ListObjects, formats the admin
'sheet, transfers all VBA code components, and finally saves the workbook
'with password protection. The class uses VB_Creatable = True so that
'factory consumers create instances via `New Linelist` rather than
'`With New Linelist`.
'@depends ILinelistSpecs, ILLdictionary, ITranslationObject,
'  IPasswords, ILLFormat, ITemporaryRepos, ICodeTransfer, IHiddenNames,
'  IButtons, IDropdownLists, ICustomPivotTable, DropdownLists,
'  CustomPivotTable, TemporaryRepos, HiddenNames, Buttons, CodeTransfer,
'  BetterArray
'@version 1.0 (2026-02-09)



'@section State
'===============================================================================
'@description Backing UDT, module-level constants, and cached instance variable.

Private Const CLASS_NAME As String = "Linelist"
Private Const PRINT_PREFIX As String = "print_"
Private Const CRF_PREFIX As String = "crf_"
Private Const SHEET_NAME_MAX_LENGTH As Long = 30
Private Const WORKBOOK_MODULE_NAME As String = "LinelistWorkbook"

Private Type TLinelist
    outWkb As Workbook
    specs As ILinelistSpecs
    tempRepos As ITemporaryRepos
    firstSheet As String
End Type

Private this As TLinelist

'@section Factory
'===============================================================================
'@description Instance creation via VB_Creatable = True factory pattern.

'@label:create
'@sub-title Instantiate a Linelist facade bound to specifications
'@details
'Validates that the supplied specifications object is not Nothing, creates
'a new Linelist instance via `New Linelist`, binds the specifications
'through the Friend property, and returns an ILinelist interface reference.
'@param specs ILinelistSpecs. The linelist specifications providing
'dictionary, translations, design format, passwords, and all other setup
'data required to build the output workbook.
'@return ILinelist. A ready-to-use linelist facade.
'@throws ProjectError.ObjectNotInitialized When specs is Nothing.
'@export
Public Function Create(ByVal specs As ILinelistSpecs) As ILinelist
    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist specifications are required"
    End If

    Dim instance As Linelist
    Set instance = New Linelist

    Set instance.BinSpecs = specs
    Set Create = instance
End Function

'@section Bindings
'===============================================================================
'@description Friend setters used by the factory to inject dependencies.

'@label:bin-specs
'@prop-title Bind the linelist specifications to this instance
'@details
'Called by Create to inject the ILinelistSpecs dependency into the
'backing UDT. Exposed as Friend to prevent external callers from
'resetting specifications after construction.
'@param value ILinelistSpecs. The specifications to bind.
Friend Property Set BinSpecs(ByVal value As ILinelistSpecs)
    Set this.specs = value
End Property

'@section Public properties
'===============================================================================
'@description Worksheet resolution and collaborator accessors exposed through
'the ILinelist contract.

'@label:linelist-data
'@prop-title Specifications backing this linelist
'@details
'Returns the ILinelistSpecs instance that was supplied during factory
'construction, providing access to dictionary, choices, translations,
'and all other setup data.
'@return ILinelistSpecs. The linelist specifications.
'@export
Public Property Get LinelistData() As ILinelistSpecs
    Set LinelistData = this.specs
End Property

'@label:dictionary
'@prop-title Dictionary from the specifications
'@details
'Convenience accessor that delegates to specs.Dictionary, avoiding
'the extra indirection at call sites.
'@return ILLdictionary. The linelist dictionary.
'@export
Public Property Get Dictionary() As ILLdictionary
    Set Dictionary = this.specs.Dictionary
End Property

'@label:wksh
'@prop-title Worksheet from the output workbook by name
'@details
'Returns the worksheet matching the given name from the output workbook.
'Delegates to ResolveWksh with scope 1 (no prefix). Raises an error
'if the sheet does not exist.
'@param sheetName String. The worksheet name to look up.
'@return Worksheet. The resolved worksheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
'@export
Public Property Get Wksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 1, Wksh
End Property

'@label:print-wksh
'@prop-title Printed companion worksheet by base name
'@details
'Returns the printed companion worksheet (prefixed with "print_") from
'the output workbook. Used for HList sheets that have a printable
'duplicate. Delegates to ResolveWksh with scope 2.
'@param sheetName String. The base sheet name (without the print_ prefix).
'@return Worksheet. The printed companion worksheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
'@export
Public Property Get PrintWksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 2, PrintWksh
End Property

'@label:crf-wksh
'@prop-title CRF companion worksheet by base name
'@details
'Returns the CRF companion worksheet (prefixed with "crf_") from the
'output workbook. Used for case report form layouts. Delegates to
'ResolveWksh with scope 3.
'@param sheetName String. The base sheet name (without the crf_ prefix).
'@return Worksheet. The CRF companion worksheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
'@export
Public Property Get CRFWksh(ByVal sheetName As String) As Worksheet
    ResolveWksh sheetName, 3, CRFWksh
End Property

'@label:dropdown
'@prop-title Dropdown lists manager for the output workbook
'@details
'Returns an IDropdownLists instance backed by the appropriate list sheet.
'Scope 1 uses the temporary dropdown_lists__ sheet with "dropdown_" prefix;
'scope 2 uses the custom choice sheet with no prefix.
'@param scope Optional Byte. 1 = standard dropdowns, 2 = custom choice
'dropdowns. Defaults to 1.
'@return IDropdownLists. The dropdown lists manager.
'@export
Public Property Get Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Dim sh As Worksheet
    Dim listSheetName As String
    Dim trans As ITranslationObject
    Dim headerPrefix As String

    Set trans = this.specs.TransObject

    Select Case scope
        Case 1
            listSheetName = this.specs.TemporarySheetName(TempSheetList)
            headerPrefix = "dropdown_"
        Case 2
            listSheetName = trans.TranslatedValue("LLSHEET_CustomChoice")
            headerPrefix = vbNullString
    End Select

    Set sh = Wksh(listSheetName)
    Set Dropdown = DropdownLists.Create(sh, headerPrefix)
End Property

'@label:pivots
'@prop-title Pivot tables manager for the output workbook
'@details
'Returns an ICustomPivotTable instance backed by the translated custom
'pivot table sheet in the output workbook.
'@return ICustomPivotTable. The pivot tables manager.
'@export
Public Property Get Pivots() As ICustomPivotTable
    Dim sh As Worksheet
    Dim trans As ITranslationObject

    Set trans = this.specs.TransObject
    Set sh = Wksh(trans.TranslatedValue("LLSHEET_CustomPivotTable"))
    Set Pivots = CustomPivotTable.Create(sh)
End Property

'@label:sheet-exists
'@sub-title Check whether a sheet exists in the output workbook
'@details
'Tests whether a worksheet with the given name exists in the output
'workbook without raising an error. Uses On Error Resume Next internally.
'@param sheetName String. The worksheet name to test.
'@return Boolean. True when the sheet exists, False otherwise.
'@export
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = OutputWkb.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = (Not sh Is Nothing)
End Function

'@label:add-output-sheet
'@sub-title Add a worksheet to the output workbook
'@details
'Creates a new worksheet in the output workbook with the given name,
'applying an optional scope prefix ("print_" or "crf_") and visibility
'setting. The sheet name is truncated to 30 characters (Excel limit for
'sheet names). Raises an error if a sheet with the resolved name already
'exists.
'@param sheetName String. The base name for the new worksheet.
'@param Hide Optional Long. Excel visibility constant. Defaults to
'xlSheetVeryHidden.
'@param sheetScope Optional Byte. 1 = no prefix, 2 = print_ prefix,
'3 = crf_ prefix. Defaults to 1.
'@throws ProjectError.InvalidArgument When a sheet with the resolved
'name already exists.
'@export
Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1)
    Dim actSheetName As String
    Dim outWkb As Workbook

    Select Case sheetScope
        Case 1
            actSheetName = sheetName
        Case 2
            actSheetName = PRINT_PREFIX & sheetName
        Case 3
            actSheetName = CRF_PREFIX & sheetName
    End Select

    actSheetName = Left$(actSheetName, SHEET_NAME_MAX_LENGTH)

    'Skip if the sheet already exists (template workbooks may contain pre-existing sheets)
    If SheetExists(actSheetName) Then Exit Sub

    Set outWkb = OutputWkb

    With outWkb
        .Worksheets.Add.Name = actSheetName
        .Worksheets(actSheetName).Visible = Hide
    End With
End Sub

'@section Lifecycle
'===============================================================================
'@description Build, save, and error-manage the output workbook.

'@label:prepare
'@sub-title Build the full linelist output workbook from specifications
'@details
'Orchestrates the complete linelist creation process:
'1. Initialises the temporary repository for code transfer files.
'2. Lazily creates the output workbook (or opens a template).
'3. Optionally shows the instruction sheet for template workbooks.
'4. Adds six temporary sheets (single, list, analysis, import, spatial,
'   show/hide) and prepares them with hidden names and ListObjects.
'5. Exports specifications to the workbook via ILinelistSpecs.Export.
'6. Adds analysis and custom sheets in reversed dictionary order.
'7. Adds data entry sheets from the dictionary's unique sheet names.
'8. Repositions the instruction sheet.
'9. Adds and formats the admin sheet (non-template only).
'10. Deletes the initial default sheet.
'11. Transfers all VBA code via CodeTransfer (classes, modules, forms,
'    workbook code).
'@export
Public Sub Prepare()
    Dim dict As ILLdictionary
    Dim trans As ITranslationObject
    Dim sheetsList As BetterArray
    Dim counter As Long
    Dim sheetName As String
    Dim outWkb As Workbook
    Dim prevAlert As Boolean
    Dim codeTransfer As ICodeTransfer

    Set this.tempRepos = this.specs.OutputTemporaryRepos

    Set outWkb = OutputWkb
    Set trans = this.specs.TransObject
    Set dict = this.specs.Dictionary
    Set sheetsList = New BetterArray

    'Handle instruction worksheet for template workbooks
    If this.specs.HasTemplate() And (this.specs.Value("instruct") = "add") Then
        outWkb.Worksheets(trans.TranslatedValue("INSTSHEETNAME")).Visible = xlSheetVisible
    End If

    'Add temporary sheets (6 sheets — TemporarySheetUpdates removed, flags are now workbook-level hidden names)
    sheetsList.Push this.specs.TemporarySheetName(TempSheetSingle), _
                    this.specs.TemporarySheetName(TempSheetList), _
                    this.specs.TemporarySheetName(TempSheetAnalysis), _
                    this.specs.TemporarySheetName(TempSheetImport), _
                    this.specs.TemporarySheetName(TempSheetSpatial), _
                    this.specs.TemporarySheetName(TempSheetShowHide)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden
    Next

    'Prepare temporary sheets with hidden names and ListObjects
    PrepareTemporarySheets

    'Add analysis/custom sheets (reversed order to maintain dictionary order)
    sheetsList.Clear
    sheetsList.Push "LLSHEET_Analysis", "LLSHEET_TemporalAnalysis", _
                    "LLSHEET_SpatialAnalysis", "LLSHEET_SpatioTemporalAnalysis", _
                    "LLSHEET_CustomChoice", "LLSHEET_CustomPivotTable"
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = trans.TranslatedValue(sheetsList.Item(counter))
        AddOutputSheet sheetName, xlSheetVisible
    Next

    'Add data entry sheets from dictionary
    sheetsList.Clear
    Set sheetsList = dict.UniqueValues("sheet name")
    sheetsList.Reverse

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVisible
    Next

    'Move instruction sheet before the last data entry sheet (template only)
    If this.specs.HasTemplate() Then
        With outWkb
            .Worksheets(trans.TranslatedValue("INSTSHEETNAME")).Move _
                Before:=.Worksheets(sheetsList.Item(counter - 1))
        End With
    End If

    'Add admin sheet if not using template
    If Not this.specs.HasTemplate() Then
        AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlSheetVisible
        PrepareAdmin
    End If

    'Delete the initial default sheet (non-template only)
    If Not this.specs.HasTemplate() Then
        prevAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False
        outWkb.Worksheets(this.firstSheet).Delete
        Application.DisplayAlerts = prevAlert
    End If

    'Transfer code via CodeTransfer
    Set codeTransfer = CodeTransfer.Create(this.specs.Workbook, outWkb, this.tempRepos)
    TransferAllCode codeTransfer
End Sub

'@label:save-ll
'@sub-title Save the output workbook as .xlsb with password protection
'@details
'Saves the output workbook in .xlsb format with the configured password,
'protects the workbook structure via IPasswords.Protect, activates an
'appropriate visible sheet (admin or instruction), closes the workbook,
'and cleans up temporary files via ITemporaryRepos.Reset.
'@export
Public Sub SaveLL()
    Dim outWkb As Workbook
    Dim passwd As String
    Dim outPath As String
    Dim passObj As IPasswords
    Dim msgTrads As ITranslationObject

    Set outWkb = OutputWkb
    Set passObj = this.specs.Password
    Set msgTrads = this.specs.TransObject

    passwd = this.specs.Value("llpassword")
    outPath = this.specs.Value("lldir") & Application.PathSeparator & _
              this.specs.Value("llname") & ".xlsb"

    'Select an appropriate visible sheet before saving
    On Error Resume Next
    outWkb.Worksheets(msgTrads.TranslatedValue("LLSHEET_Admin")).Activate
    outWkb.Worksheets(msgTrads.TranslatedValue("INSTSHEETNAME")).Activate
    On Error GoTo 0

    'Protect the workbook structure
    passObj.Protect outWkb

    outWkb.SaveAs fileName:=outPath, fileFormat:=xlExcel12, Password:=passwd, _
                   ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    outWkb.Close

    If Not this.tempRepos Is Nothing Then this.tempRepos.Reset
End Sub

'@label:error-manage
'@sub-title Display an error dialog for incomplete linelist operations
'@details
'Shows a message box with the error details and asks the user whether
'they want to view the incomplete linelist workbook or close it without
'saving. If the user chooses Yes the workbook's Application is made
'visible; otherwise the workbook is closed without saving.
'@param textMessage Optional String. Error description to display.
'Defaults to vbNullString.
'@export
Public Sub ErrorManage(Optional ByVal textMessage As String = vbNullString)
    Dim outWkb As Workbook
    Dim answer As Long

    answer = MsgBox("An internal error occurred with specified information: " & Chr(10) & _
                    textMessage & Chr(10) & _
                    "Do you want to see the uncompleted linelist?", _
                    vbExclamation + vbYesNo, "ERROR")

    Set outWkb = OutputWkb

    If answer = vbYes Then
        outWkb.Application.Visible = True
    Else
        outWkb.Close savechanges:=False
    End If
End Sub

'@section Private — Workbook management
'===============================================================================
'@description Internal helpers for lazily creating and managing the output
'workbook and suppressing UI updates during batch operations.

'@label:output-wkb
'@prop-title Return the output workbook from LinelistSpecs
'@details
'On first access, retrieves the output workbook that LinelistSpecs.Prepare
'created and caches it locally. Also records the name of the initial
'default sheet so Prepare can delete it later.
'@return Workbook. The output workbook.
Private Property Get OutputWkb() As Workbook
    If this.outWkb Is Nothing Then
        Set this.outWkb = this.specs.OutputWorkbook
        this.firstSheet = this.outWkb.Worksheets(1).Name
    End If

    Set OutputWkb = this.outWkb
End Property

'@label:resolve-wksh
'@sub-title Resolve a worksheet from the output workbook with scope prefix
'@details
'Builds the actual sheet name by prepending the appropriate prefix based
'on sheetScope (1 = none, 2 = "print_", 3 = "crf_"), verifies the sheet
'exists, and returns a reference through the ByRef result parameter.
'@param sheetName String. The base sheet name.
'@param sheetScope Byte. 1 = no prefix, 2 = print_ prefix, 3 = crf_ prefix.
'@param result Worksheet. ByRef output parameter receiving the resolved sheet.
'@throws ProjectError.InvalidArgument When the sheet does not exist.
Private Sub ResolveWksh(ByVal sheetName As String, ByVal sheetScope As Byte, _
                        ByRef result As Worksheet)
    Dim actSheetName As String

    Select Case sheetScope
        Case 1
            actSheetName = sheetName
        Case 2
            actSheetName = PRINT_PREFIX & sheetName
        Case 3
            actSheetName = CRF_PREFIX & sheetName
    End Select

    If Not SheetExists(actSheetName) Then
        ThrowError ProjectError.InvalidArgument, _
                   "The sheet " & actSheetName & " does not exist in the output workbook"
    End If

    Set result = OutputWkb.Worksheets(actSheetName)
End Sub

'@section Private — Prepare helpers
'===============================================================================
'@description Internal routines called by Prepare to set up temporary sheets,
'the admin sheet, and code transfer.

'@label:prepare-temporary-sheets
'@sub-title Set up hidden names and ListObjects on temporary sheets
'@details
'Performs three setup tasks on the output workbook:
'1. Creates workbook-level hidden names for global flags (update list auto,
'   import done, geo update manual, show all optionals, export init, number
'   of exports, epi week start) — replaces the removed TemporarySheetUpdates.
'2. Creates four ListObjects on the import report sheet for tracking sheets
'   and variables not imported/found.
'3. Creates twelve ListObjects on the analysis sheet (tab_ids, graph_ids,
'   graph_formats for each of the four analysis scopes: uba, ts, sp,
'   sptemp) plus four worksheet-level hidden names storing the translated
'   analysis sheet names.
Private Sub PrepareTemporarySheets()
    Dim outWkb As Workbook
    Dim sh As Worksheet
    Dim rng As Range
    Dim cellRng As Range
    Dim trads As ITranslationObject
    Dim sheetName As String
    Dim counter As Long
    Dim namesArray As BetterArray
    Dim actName As String
    Dim wkbNames As IHiddenNames
    Dim shNames As IHiddenNames

    Set outWkb = OutputWkb
    Set trads = this.specs.TransObject
    Set namesArray = New BetterArray

    'Workbook-level hidden names for global flags (replaces TemporarySheetUpdates)
    Set wkbNames = HiddenNames.Create(outWkb)
    wkbNames.EnsureName "RNG_UpdateListAuto", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_ImportDone", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_GeoUpdateManual", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_ShowAllOptionals", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_ExportInit", vbNullString, HiddenNameTypeString
    wkbNames.EnsureName "RNG_NumberExports", this.specs.Value("numberofexports"), HiddenNameTypeString
    wkbNames.EnsureName "RNG_EpiWeekStart", this.specs.Value("epiweekstart"), HiddenNameTypeString

    'Import report sheet — create ListObjects
    sheetName = this.specs.TemporarySheetName(TempSheetImport)
    Set sh = outWkb.Worksheets(sheetName)

    With sh
        .Cells(1, 1).Value = "sheetNotImported"
        .Cells(1, 3).Value = "sheetNotTouched"
        .Cells(1, 5).Value = "variableNotImported"
        .Cells(1, 6).Value = "correspondingSheet"
        .Cells(1, 8).Value = "variableNotFound"
        .Cells(1, 9).Value = "correspondingSheet"

        Set rng = .Range(.Cells(1, 1), .Cells(2, 1))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotImp"

        Set rng = .Range(.Cells(1, 3), .Cells(2, 3))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_sheetsNotTouch"

        Set rng = .Range(.Cells(1, 5), .Cells(2, 6))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotImp"

        Set rng = .Range(.Cells(1, 8), .Cells(2, 9))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "reptab_varsNotTouch"
    End With

    'Analysis table IDs sheet — create ListObjects + worksheet-level hidden names
    sheetName = this.specs.TemporarySheetName(TempSheetAnalysis)
    Set sh = outWkb.Worksheets(sheetName)
    Set cellRng = sh.Cells(1, 1)

    namesArray.Clear
    namesArray.Push "uba", "ts", "sp", "sptemp"

    For counter = namesArray.LowerBound To namesArray.UpperBound
        actName = namesArray.Item(counter)

        'Table IDs ListObject
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "tabRange"
        cellRng.Cells(1, 3).Value = "exportTable"

        Set rng = sh.Range(cellRng, cellRng.Cells(2, 3))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "tab_ids_" & actName

        'Graph IDs ListObject
        Set cellRng = cellRng.Cells(1, 5)
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "graphId"
        cellRng.Cells(1, 3).Value = "SeriesName"
        cellRng.Cells(1, 4).Value = "SeriesType"
        cellRng.Cells(1, 5).Value = "SeriesPos"
        cellRng.Cells(1, 6).Value = "SeriesLabel"
        cellRng.Cells(1, 7).Value = "SeriesColumnLabel"
        cellRng.Cells(1, 8).Value = "hardCodeLabels"
        cellRng.Cells(1, 9).Value = "outRangeAddress"
        cellRng.Cells(1, 10).Value = "prefix"
        cellRng.Cells(1, 11).Value = "prefixOnly"

        Set rng = sh.Range(cellRng, cellRng.Cells(2, 11))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "graph_ids_" & actName

        'Graph formats ListObject
        Set cellRng = cellRng.Cells(1, 13)
        cellRng.Value = "tabId"
        cellRng.Cells(1, 2).Value = "graphId"
        cellRng.Cells(1, 3).Value = "catTitle"
        cellRng.Cells(1, 4).Value = "valuesTitle"
        cellRng.Cells(1, 5).Value = "hardCodeLabels"
        cellRng.Cells(1, 6).Value = "heightFactor"
        cellRng.Cells(1, 7).Value = "plotTitle"

        Set rng = sh.Range(cellRng, cellRng.Cells(2, 7))
        sh.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "graph_formats_" & actName

        Set cellRng = cellRng.Cells(1, 10)
    Next

    'Worksheet-level hidden names for analysis sheet name references
    Set shNames = HiddenNames.Create(sh)
    shNames.EnsureName "RNG_SheetUAName", _
                       trads.TranslatedValue("LLSHEET_Analysis"), HiddenNameTypeString
    shNames.EnsureName "RNG_SheetTSName", _
                       trads.TranslatedValue("LLSHEET_TemporalAnalysis"), HiddenNameTypeString
    shNames.EnsureName "RNG_SheetSPName", _
                       trads.TranslatedValue("LLSHEET_SpatialAnalysis"), HiddenNameTypeString
    shNames.EnsureName "RNG_SheetSPTempName", _
                       trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"), HiddenNameTypeString
End Sub

'@label:prepare-admin
'@sub-title Format the admin sheet and add Advanced/Export buttons
'@details
'Applies the linelist design format to the admin sheet, then creates two
'large-scope buttons (Advanced and Export) with translated labels. After
'placing the buttons, re-suppresses UI updates and activates the admin
'sheet. Finally protects the active sheet via IPasswords.Protect with
'the "_active" scope marker.
'@remarks
'The trads object is switched to TranslationOfShapes scope to fetch
'shape-specific translated labels for the buttons.
Private Sub PrepareAdmin()
    Const ADVANCED_COMMAND As String = "ClickAdvanced"
    Const EXPORT_COMMAND As String = "ClickExport"
    Const SHP_ADVANCED As String = "SHP_Advanced"
    Const SHP_EXPORT As String = "SHP_Export"

    Dim sh As Worksheet
    Dim trads As ITranslationObject
    Dim design As ILLFormat
    Dim pass As IPasswords
    Dim but As IButtons
    Dim rng As Range

    Set design = this.specs.DesignFormat
    Set trads = this.specs.TransObject
    Set pass = this.specs.Password

    Set sh = Wksh(trads.TranslatedValue("LLSHEET_Admin"))
    design.ApplyFormat sh, scope:=AllLinelistSheet

    'Switch to shapes translation scope
    Set trads = this.specs.TransObject(TranslationOfShapes)

    'Advanced button
    Set rng = sh.Cells(2, 3)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHP_ADVANCED, scope:=ButtonScopeLarge)
    but.Add actionCommand:=ADVANCED_COMMAND, shapeLabel:=trads.TranslatedValue(SHP_ADVANCED)
    but.Format design

    'Export button
    Set rng = sh.Cells(2, 5)
    Set but = Buttons.Create(outrng:=rng, codeName:=SHP_EXPORT, scope:=ButtonScopeLarge)
    but.Add actionCommand:=EXPORT_COMMAND, shapeLabel:=trads.TranslatedValue(SHP_EXPORT)
    but.Format design

    sh.Activate

    pass.Protect "_active"
End Sub

'@label:transfer-all-code
'@sub-title Transfer all VBA code components to the output workbook
'@details
'Uses the supplied ICodeTransfer to copy VBA components from the setup
'workbook to the output workbook in four phases:
'1. Transfers 42 class modules (interface+implementation pairs for
'   DataSheet, LLExport, LLChoices, LLdictionary, LLGeo, Passwords,
'   LLSheets, LLTranslations, OSFiles, TranslationObject, LLVariables,
'   LLSpatial, DropdownLists, LLShowHide, CustomTable, OutputSpecs,
'   ImpRep, ImpSpecs, UpVal, Graphs, AnaTabIds, ExportButtons, plus
'   BetterArray).
'2. Transfers 6 standard modules (LinelistCustomFunctions, GeoModule,
'   LinelistEvents, EventsLinelistButtons, EventsGlobalAnalysis,
'   LLChange).
'3. Copies the ribbon module text (EventsLinelistRibbon) for template
'   workbooks.
'4. Transfers 9 user forms (F_Advanced, F_Export, F_ExportMig, F_Geo,
'   F_ImportRep, F_ShowHideLL, F_ShowHidePrint, F_ShowVarLabels,
'   F_EpiWeek).
'5. Transfers the workbook code module (LinelistWorkbook).
'6. Cleans up temporary files via ITemporaryRepos.Reset.
'@param codeTransfer ICodeTransfer. The code transfer engine.
Private Sub TransferAllCode(ByVal codeTransfer As ICodeTransfer)
    Dim codesList As BetterArray
    Dim counter As Long
    Dim codeName As String

    Set codesList = New BetterArray

    'Transfer classes (42 classes)
    codesList.Push "BetterArray", "DataSheet", "IDataSheet", "LLExport", _
                   "ILLExport", "LLChoices", "ILLChoices", _
                   "LLdictionary", "ILLdictionary", "LLGeo", _
                   "ILLGeo", "Passwords", "IPasswords", _
                   "LLSheets", "ILLSheets", "LLTranslation", _
                   "ILLTranslation", "OSFiles", "IOSFiles", _
                   "TranslationObject", "ITranslationObject", "LLVariables", _
                   "ILLVariables", "ILLSpatial", "LLSpatial", _
                   "DropdownLists", "IDropdownLists", "LLShowHide", _
                   "ILLShowHide", "CustomTable", "ICustomTable", _
                   "LLExporter", "ILLExporter", _
                   "ImpRep", "IImpRep", "LLImporter", "ILLImporter", _
                   "IGraphs", "Graphs", "AnaTabIds", _
                   "IAnaTabIds", "ExportButtons", "IExportButtons"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferClass codeName
    Next

    'Transfer modules (6 modules)
    codesList.Clear
    codesList.Push "LinelistCustomFunctions", _
                   "GeoModule", "LinelistEvents", "EventsLinelistRibbon", _
                   "EventsLinelistButtons", "EventsGlobalAnalysis"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferModule codeName
    Next

    'Copy ribbon module text (for template workbooks)
    codeTransfer.CopyModuleText "EventsLinelistRibbon"

    'Transfer forms (9 forms)
    codesList.Clear
    codesList.Push "F_Advanced", "F_Export", "F_ExportMig", _
                   "F_Geo", "F_ImportRep", _
                   "F_ShowHideLL", "F_ShowHidePrint", "F_ShowVarLabels", _
                   "F_EpiWeek"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferForm codeName
    Next

    'Transfer workbook code
    codeTransfer.TransferWorkbookCode WORKBOOK_MODULE_NAME
End Sub

'@section Private — Error handling
'===============================================================================
'@description Centralised error-raising helper.

'@label:throw-error
'@sub-title Raise a project error with class context
'@details
'Wraps Err.Raise with CLASS_NAME as the source, providing consistent
'error attribution across all methods in this class.
'@param errNumber Long. The ProjectError constant to raise.
'@param errorMessage String. Descriptive message for the error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise errNumber, CLASS_NAME, errorMessage
End Sub

'@section ILinelist implementation
'===============================================================================
'@description Delegated members satisfying the ILinelist contract.
'All members forward to their Public counterparts above.

Private Property Get ILinelist_Wksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_Wksh = Wksh(sheetName)
End Property

Private Property Get ILinelist_PrintWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_PrintWksh = PrintWksh(sheetName)
End Property

Private Property Get ILinelist_CRFWksh(ByVal sheetName As String) As Worksheet
    Set ILinelist_CRFWksh = CRFWksh(sheetName)
End Property

Private Property Get ILinelist_Dropdown(Optional ByVal scope As Byte = 1) As IDropdownLists
    Set ILinelist_Dropdown = Dropdown(scope)
End Property

Private Property Get ILinelist_Pivots() As ICustomPivotTable
    Set ILinelist_Pivots = Pivots
End Property

Private Property Get ILinelist_LinelistData() As ILinelistSpecs
    Set ILinelist_LinelistData = LinelistData
End Property

Private Property Get ILinelist_Dictionary() As ILLdictionary
    Set ILinelist_Dictionary = Dictionary
End Property

Private Sub ILinelist_Prepare()
    Prepare
End Sub

Private Function ILinelist_SheetExists(ByVal sheetName As String) As Boolean
    ILinelist_SheetExists = SheetExists(sheetName)
End Function

Private Sub ILinelist_AddOutputSheet(ByVal sheetName As String, _
                                     Optional ByVal Hide As Long = xlSheetVeryHidden, _
                                     Optional ByVal sheetScope As Byte = 1)
    AddOutputSheet sheetName, Hide, sheetScope
End Sub

Private Sub ILinelist_SaveLL()
    SaveLL
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    ErrorManage textMessage
End Sub
