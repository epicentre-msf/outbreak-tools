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
Private Const WORKBOOK_MODULE_NAME As String = "EventLinelistWorkbook"
Private Const SHEET_GEO As String = "Geo"

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
'sheet names). When beforeSheet is supplied and that sheet exists in
'the output workbook, the new sheet is inserted immediately before it;
'otherwise the sheet is appended at the default position.
'@param sheetName String. The base name for the new worksheet.
'@param Hide Optional Long. Excel visibility constant. Defaults to
'xlSheetVeryHidden.
'@param sheetScope Optional Byte. 1 = no prefix, 2 = print_ prefix,
'3 = crf_ prefix. Defaults to 1.
'@param beforeSheet Optional String. Name of an existing worksheet to
'insert before. Defaults to vbNullString (append at default position).
'@export
Public Sub AddOutputSheet(ByVal sheetName As String, _
                          Optional ByVal Hide As Long = xlSheetVeryHidden, _
                          Optional ByVal sheetScope As Byte = 1, _
                          Optional ByVal beforeSheet As String = vbNullString)
    Dim actSheetName As String
    Dim outWkb As Workbook
    Dim anchorSh As Worksheet

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

    'Resolve the anchor sheet if provided
    If LenB(beforeSheet) > 0 Then
        On Error Resume Next
        Set anchorSh = outWkb.Worksheets(beforeSheet)
        On Error GoTo 0
    End If

    With outWkb
        If Not anchorSh Is Nothing Then
            .Worksheets.Add(Before:=anchorSh).Name = actSheetName
        Else
            .Worksheets.Add.Name = actSheetName
        End If
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
    Dim codeTrans As ICodeTransfer

    If Not (this.specs.Prepared) Then
        ThrowError ProjectError.ErrorUnexpectedState, _ 
         "Linelist specifications need to be prepared before Linelist creation"
    End If

    Set this.tempRepos = this.specs.OutputTemporaryRepos

    Set outWkb = OutputWkb
    Set dict = this.specs.Dictionary
    Set sheetsList = New BetterArray

    InitialiseWorkbookNames
    InitialiseTranslationNames

    Set trans = this.specs.TransObject
    'Add data entry sheets from dictionary before Geo
    sheetsList.Clear
    Set sheetsList = dict.UniqueValues("sheet name")

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVisible, beforeSheet:=SHEET_GEO
    Next

    'Add analysis/custom sheets before Geo (reversed order to maintain dictionary order)
    sheetsList.Clear
    sheetsList.Push trans.TranslatedValue("LLSHEET_Analysis"), _ 
                    trans.TranslatedValue("LLSHEET_TemporalAnalysis"), _
                    trans.TranslatedValue("LLSHEET_SpatialAnalysis"), _ 
                    trans.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"), _
                    trans.TranslatedValue("LLSHEET_CustomChoice"), _ 
                    trans.TranslatedValue("LLSHEET_CustomPivotTable")

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVisible, beforeSheet:=SHEET_GEO
    Next


    'Add temporary sheets before the Geo worksheet
    sheetsList.Clear
    '(6 sheets — TemporarySheetUpdates removed, flags are now workbook-level hidden names)
    sheetsList.Push this.specs.TemporarySheetName(TempSheetSingle), _
                    this.specs.TemporarySheetName(TempSheetList), _
                    this.specs.TemporarySheetName(TempSheetAnalysis), _
                    this.specs.TemporarySheetName(TempSheetImport), _
                    this.specs.TemporarySheetName(TempSheetSpatial), _
                    this.specs.TemporarySheetName(TempSheetShowHide)

    For counter = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = sheetsList.Item(counter)
        AddOutputSheet sheetName, xlSheetVeryHidden, beforeSheet:=SHEET_GEO
    Next

    'Prepare temporary sheets with hidden names and ListObjects
    PrepareTemporarySheets

    'Initialise language and translated HiddenNames on the output workbook

    'Handle instruction sheet (template only)
    If this.specs.HasTemplate() Then
        Dim instrName As String
        Dim instrSheet As Worksheet

        instrName = trans.TranslatedValue("INSTSHEETNAME")

        On Error Resume Next
        Set instrSheet = outWkb.Worksheets(instrName)
        On Error GoTo 0

        If Not instrSheet Is Nothing Then
            'Unhide before Move (hidden sheets cannot be moved)
            instrSheet.Visible = xlSheetVisible
            'Move instruction sheet to first position
            instrSheet.Move Before:=outWkb.Worksheets(1)
            
            'Re-hide if instruct flag is not "add"
            If this.specs.Value("instruct") = "No" Then
               instrSheet.Visible = xlSheetVeryHidden
            End If
        End If
    End If

    ' 'Add admin sheet if not using template
    If Not this.specs.HasTemplate() Then
        AddOutputSheet trans.TranslatedValue("LLSHEET_Admin"), xlSheetVisible, _
                       beforeSheet:=SHEET_GEO
        PrepareAdmin
    End If

    ' 'Delete the initial default sheet (non-template only)
    If Not this.specs.HasTemplate() Then
        prevAlert = Application.DisplayAlerts
        Application.DisplayAlerts = False
        outWkb.Worksheets(this.firstSheet).Delete
        Application.DisplayAlerts = prevAlert
    End If

    'Transfer code via CodeTransfer
    Set codeTrans = CodeTransfer.Create(this.specs.DesWorkbook, outWkb, this.tempRepos)
    TransferAllCode codeTrans
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

    'Protect the workbook structure, add debugging mechanism
    passObj.Protect outWkb
    passObj.EnsureDebugExitHandler outWkb

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
        Set this.outWkb = this.specs.LLWorkbook
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
    wkbNames.SetValue "RNG_UpdateListAuto", vbNullString
    wkbNames.EnsureName "RNG_ImportDone", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_ImportDone", vbNullString
    wkbNames.EnsureName "RNG_GeoUpdateManual", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_GeoUpdateManual", vbNullString
    wkbNames.EnsureName "RNG_ShowAllOptionals", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_ShowAllOptionals", vbNullString
    wkbNames.EnsureName "RNG_ExportInit", vbNullString, HiddenNameTypeString
    wkbNames.SetValue "RNG_ExportInit", vbNullString
    wkbNames.EnsureName "RNG_NumberExports", this.specs.Value("numberofexports"), HiddenNameTypeString
    wkbNames.SetValue "RNG_NumberExports", this.specs.Value("numberofexports")
    wkbNames.EnsureName "RNG_EpiWeekStart", this.specs.Value("epiweekstart"), HiddenNameTypeString
    wkbNames.SetValue "RNG_EpiWeekStart", this.specs.Value("epiweekstart")

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
    shNames.SetValue "RNG_SheetUAName", trads.TranslatedValue("LLSHEET_Analysis")
    shNames.EnsureName "RNG_SheetTSName", _
                       trads.TranslatedValue("LLSHEET_TemporalAnalysis"), HiddenNameTypeString
    shNames.SetValue "RNG_SheetTSName", trads.TranslatedValue("LLSHEET_TemporalAnalysis")
    shNames.EnsureName "RNG_SheetSPName", _
                       trads.TranslatedValue("LLSHEET_SpatialAnalysis"), HiddenNameTypeString
    shNames.SetValue "RNG_SheetSPName", trads.TranslatedValue("LLSHEET_SpatialAnalysis")
    shNames.EnsureName "RNG_SheetSPTempName", _
                       trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis"), HiddenNameTypeString
    shNames.SetValue "RNG_SheetSPTempName", trads.TranslatedValue("LLSHEET_SpatioTemporalAnalysis")
End Sub

'@label:initialise-workbook-names
'@sub-title Set language HiddenNames on the output workbook
'@details
'Creates workbook-level HiddenNames for the dictionary language
'(RNG_DictionaryLanguage) and interface language (RNG_LLLanguageCode)
'on the output workbook. These are required by LLTranslation.TransObject
'to resolve the correct language code for translation lookups.
Private Sub InitialiseWorkbookNames()
    Dim outWkb As Workbook
    Dim wkbNames As IHiddenNames
    Dim dictLang As String
    Dim llLang As String

    Set outWkb = OutputWkb
    Set wkbNames = HiddenNames.Create(outWkb)

    dictLang = this.specs.Value("setuplang")
    llLang = Split(this.specs.Value("lllang"), "-")(0)

    'Fail early when language selections are empty
    If LenB(dictLang) = 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Dictionary language (RNG_LangSetup) is empty on the Main sheet. " & _
                   "Please select a setup language before generating."
    End If

    If LenB(llLang) = 0 Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Interface language (RNG_LLForm) is empty on the Main sheet. " & _
                   "Please select an interface language before generating."
    End If

    wkbNames.EnsureName "RNG_DictionaryLanguage", dictLang, HiddenNameTypeString
    wkbNames.SetValue "RNG_DictionaryLanguage", dictLang
    wkbNames.EnsureName "RNG_LLLanguageCode", llLang, HiddenNameTypeString
    wkbNames.SetValue "RNG_LLLanguageCode", llLang
End Sub

'@label:initialise-translation-names
'@sub-title Propagate language HiddenNames to the translation worksheet and create translated names
'@details
'Calls InitialiseSheetNames on the LLTranslation instance to store
'dictionary and interface language codes as worksheet-level HiddenNames
'on the translation sheet. Then calls InitialiseHiddenNames to create
'all RNG_* translated message/sheet names on the output workbook.
Private Sub InitialiseTranslationNames()
    Dim transObj As ILLTranslation
    Dim dictLang As String
    Dim llLang As String

    dictLang = this.specs.Value("setuplang")
    llLang = Split(this.specs.Value("lllang"), "-")(0)

    Set transObj = this.specs.Translations
    transObj.InitialiseSheetNames dictLang, llLang
    transObj.InitialiseHiddenNames OutputWkb
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
'workbook to the output workbook in five phases:
'1. Transfers 49 class modules (interface+implementation pairs for
'   DataSheet, LLExport, LLChoices, LLdictionary, LLGeo, Passwords,
'   LLSheets, LLTranslation, OSFiles, TranslationObject, LLVariables,
'   LLSpatial, DropdownLists, ShowHideManager, CustomTable, LLExporter,
'   LLImporter, Graphs, AnaTabIds, ExportButton, HiddenNames,
'   EventLinelist, ApplicationState, plus BetterArray and Checking).
'2. Transfers 7 standard modules (LinelistCustomFunctions, GeoModule,
'   LinelistEventsManager, EventsLinelistButtons, FormLogicAdvanced,
'   FormLogicExportMig, FormLogicExport). For template workbooks,
'   copies the ribbon module text (EventsLinelistRibbon); otherwise
'   transfers it as a new module.
'3. Transfers 9 user forms (F_Advanced, F_Export, F_ExportMig, F_Geo,
'   F_ImportRep, F_ShowHideLL, F_ShowHidePrint, F_ShowVarLabels,
'   F_EpiWeek).
'4. Transfers the workbook code module (EventLinelistWorkbook).
'@param codeTransfer ICodeTransfer. The code transfer engine.
Private Sub TransferAllCode(ByVal codeTransfer As ICodeTransfer)
    Dim codesList As BetterArray
    Dim counter As Long
    Dim codeName As String

    Set codesList = New BetterArray

    'Transfer classes (49 classes)
    codesList.Push "BetterArray", "DataSheet", "IDataSheet", "LLExport", _
                   "ILLExport", "LLChoices", "ILLChoices", _
                   "LLdictionary", "ILLdictionary", "LLGeo", _
                   "ILLGeo", "Passwords", "IPasswords", _
                   "IChecking", "Checking", _
                   "LLSheets", "ILLSheets", "LLTranslation", _
                   "ILLTranslation", "OSFiles", "IOSFiles", _
                   "TranslationObject", "ITranslationObject", "LLVariables", _
                   "ILLVariables", "ILLSpatial", "LLSpatial", _
                   "DropdownLists", "IDropdownLists", "ShowHideManager", _
                   "IShowHideManager", "CustomTable", "ICustomTable", _
                   "LLExporter", "ILLExporter", _
                    "LLImporter", "ILLImporter", _
                   "IGraphs", "Graphs", "AnaTabIds", _
                   "IAnaTabIds", "ExportButton", "IExportButton", _
                   "HiddenNames", "IHiddenNames", _
                   "EventLinelist", "IEventLinelist", _
                   "ApplicationState", "IApplicationState"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferClass codeName
    Next

    'Transfer modules (7 modules + conditional ribbon)
    codesList.Clear
    codesList.Push "LinelistCustomFunctions", _
                   "GeoModule", "LinelistEventsManager", _
                   "EventsLinelistButtons", _
                   "FormLogicAdvanced", "FormLogicExportMig", _
                   "FormLogicExport"

    For counter = codesList.LowerBound To codesList.UpperBound
        codeName = codesList.Item(counter)
        codeTransfer.TransferModule codeName
    Next

    'Transfer or copy ribbon module depending on template status
    If this.specs.HasTemplate() Then
        'Template already contains EventsLinelistRibbon; overwrite its code
        codeTransfer.CopyModuleText "EventsLinelistRibbon"
    Else
        'Non-template: import EventsLinelistRibbon as a new module
        codeTransfer.TransferModule "EventsLinelistRibbon"
    End If

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
                                     Optional ByVal sheetScope As Byte = 1, _
                                     Optional ByVal beforeSheet As String = vbNullString)
    AddOutputSheet sheetName, Hide, sheetScope, beforeSheet
End Sub

Private Sub ILinelist_SaveLL()
    SaveLL
End Sub

Private Sub ILinelist_ErrorManage(Optional ByVal textMessage As String = vbNullString)
    ErrorManage textMessage
End Sub
