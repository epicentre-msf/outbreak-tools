Attribute VB_Name = "EventsDesignerAdvanced"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Non-core ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, DesignerEntry, RibbonDev, LLGeo, ApplicationState, OSFiles, HiddenNames, BetterArray, DropdownLists, DropdownLists, InitTransfer, LinelistSpecs, LinelistSpecs, Linelist, Linelist, LLDataEntry, LLSheets, GenerationReport
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Non-core ribbon logics are callbacks whose absence will not fire a
'warning at workbook opening on the designer. They only execute in
'response to explicit user actions (onAction), never at ribbon load
'time (getLabel, getPressed, getVisible).

Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const PROMPT_TITLE As String = "Designer"

Private Const SHEET_TRANSLATIONS As String = "Translations"

'HiddenName storing semicolon-separated language list on the Translations sheet
Private Const SETUP_LANGUAGES_TAG As String = "__SetupTranslationsLanguages__"

'Dropdown name used by DesignerPreparation for setup languages
Private Const DROP_SETUP_LANGUAGES As String = "__setup_languages"


'@section Dev group callbacks
'===============================================================================

'@Description("Initialise the designer workbook: import translations, hide sheets, seed flags.")
'@EntryPoint
Public Sub clickDesignerInitialize(ByRef control As IRibbonControl)
    Dim prep As DesignerPreparation
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set prep = DesignerPreparation.Create(ThisWorkbook)
    prep.Prepare RibbonDev.EnsureDevelopment()

    appScope.Restore
    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickDesignerInitialize: "; errNumber; errDesc
        MsgBox "Unable to initialise designer: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Manage group callbacks
'===============================================================================

'@Description("Clear all geobase data from the Geo worksheet.")
'@EntryPoint
Public Sub clickDelGeo(ByRef control As IRibbonControl)
    Dim geoSheet As Worksheet
    Dim geo As LLGeo
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set geoSheet = ThisWorkbook.Worksheets(SHEET_GEO)
    Set geo = LLGeo.Create(geoSheet)
    geo.Clear

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickDelGeo: "; errNumber; errDesc
        MsgBox "Unable to clear geobase: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Clear all entry input ranges on the Main sheet.")
'@EntryPoint
Public Sub clickClearEnt(ByRef control As IRibbonControl)
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))
    entry.Clear

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickClearEnt: "; errNumber; errDesc
        MsgBox "Unable to clear entries: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section File and folder loading callbacks
'===============================================================================

'@Description("Load a setup file (dictionary): store path, extract languages, update dropdown.")
'@EntryPoint
Public Sub clickLoadFileDic()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState
    Dim setupBook As Workbook
    Dim tradSheet As Worksheet

    'Show the file dialog before entering busy state (dialog needs UI)
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb;*.xlsx"

    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))

    'Open the selected setup workbook read-only
    Set setupBook = Workbooks.Open(io.File(), ReadOnly:=True)

    'Validate that the setup has a Translations worksheet
    On Error Resume Next
    Set tradSheet = setupBook.Worksheets(SHEET_TRANSLATIONS)
    On Error GoTo Cleanup

    If tradSheet Is Nothing Then
        setupBook.Close saveChanges:=False
        Set setupBook = Nothing
        entry.AddInfo entry.TranslateMessage("MSG_OpeAnnule"), "edition"
        GoTo Cleanup
    End If

    'Write the setup path to the Main sheet
    entry.AddInfo io.File(), "setuppath"
    entry.AddInfo entry.TranslateMessage("MSG_ChemFich"), "edition"

    'Extract languages from the setup Translations worksheet HiddenNames
    'and update the setup languages dropdown for the designer
    ExtractAndUpdateLanguages tradSheet

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'Close the setup workbook if still open
    If Not setupBook Is Nothing Then
        setupBook.Close saveChanges:=False
    End If
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLoadFileDic: "; errNumber; errDesc
        MsgBox "Unable to load setup file: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Load a geobase file path into the Main sheet.")
'@EntryPoint
Public Sub clickLoadGeoFile()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    'Show the file dialog before entering busy state
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"

    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))
    entry.AddInfo io.File(), "geopath"

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLoadGeoFile: "; errNumber; errDesc
        MsgBox "Unable to load geobase: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Select a folder for linelist output directory.")
'@EntryPoint
Public Sub clickLinelistDir()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    'Show the folder dialog before entering busy state
    Set io = OSFiles.Create()
    io.LoadFolder

    If Not io.HasValidFolder() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))
    entry.AddInfo io.Folder(), "lldir"

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLinelistDir: "; errNumber; errDesc
        MsgBox "Unable to set linelist directory: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub

'@Description("Load a template file for linelist creation.")
'@EntryPoint
Public Sub clickLoadTemplate()
    Dim io As OSFiles
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState

    'Show the file dialog before entering busy state
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb"

    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlWait

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))
    entry.AddInfo io.File(), "temppath"
    entry.AddInfo entry.TranslateMessage("MSG_ChemFich"), "edition"

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickLoadTemplate: "; errNumber; errDesc
        MsgBox "Unable to load template: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Generation callbacks
'===============================================================================

'@Description("Import setup, prepare specifications, build output linelist workbook, and save.")
'@EntryPoint
Public Sub clickGenerate()
    Dim entry As DesignerEntry
    Dim appScope As ApplicationState
    Dim specs As LinelistSpecs
    Dim ll As Linelist
    Dim setupPath As String
    Dim sheetLists As BetterArray
    Dim counter As Long
    Dim anaOut As AnalysisOutput

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlNorthWestArrow

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))

    'Run readiness checks; exit silently when requirements are not met
    If Not ValidateGenerationReadiness(entry) Then GoTo Cleanup

    setupPath = entry.ValueOf("setuppath")

    'Build the linelist: Prepare creates the output workbook and hands it to
    'InitTransfer, which fills it from the setup file and from this designer.
    entry.AddInfo entry.TranslateMessage("MSG_ReadSetup"), "edition"

    'Initialise the generation report on the designer __checking sheet
    GenerationReport.InitReport ThisWorkbook

    Set specs = LinelistSpecs.Create(ThisWorkbook)
    specs.Prepare setupPath

    'Flush Phase 1: specification checkings (dictionary, choices, exports, etc.)
    GenerationReport.FlushCheckings GenerationReport.HarvestSpecsCheckings(specs)

    'After the preparation step of the specifications, internal specifications
    'object shift focus from the designer to the linelist workbook as they
    'are now exported.

    'Build the output linelist workbook (sheets, temp sheets, admin, code transfer)
    Set ll = Linelist.Create(specs)
    ll.Prepare

    'Flush Phase 1b: code transfer checkings. A component the output workbook
    'already carried was replaced by the designer's copy, and this is where the
    'report names it.
    If ll.HasCheckings Then
        Dim codeChecks As BetterArray
        Set codeChecks = New BetterArray
        codeChecks.LowerBound = 1
        codeChecks.Push ll.CheckingValues
        GenerationReport.FlushCheckings codeChecks
    End If

    'Build data entry worksheets (sections, variables, formatting). The sheet
    'name list is the one Linelist.Prepare already walked the dictionary for.
    Set sheetLists = ll.SheetNames

    If sheetLists.Length > 0 Then
        Dim listBld As LLDataEntry
        Dim sheetChecks As BetterArray
        Dim llSheetInfo As LLSheets
        Set sheetChecks = New BetterArray
        sheetChecks.LowerBound = 1

        'The shared LLSheets the linelist holds. This loop and TransferAllCode
        'each created their own over the same dictionary, so every row
        'resolution was computed twice.
        Set llSheetInfo = ll.SheetInfoManager

        For counter = sheetLists.LowerBound To sheetLists.UpperBound
            Set listBld = BuildOneSheet(llSheetInfo, ll, sheetLists.Item(counter))
            If Not listBld Is Nothing Then
                If listBld.HasCheckings Then
                    sheetChecks.Push listBld.CheckingValues
                End If
            End If
        Next

        'Flush Phase 2: per-sheet build checkings
        GenerationReport.FlushCheckings sheetChecks
    End If

    'Flush Phase 2b: shared dropdown checkings (linelist-level, not per-sheet)
    Dim dropChecks As BetterArray
    Set dropChecks = New BetterArray
    dropChecks.LowerBound = 1

    Dim dropStd As DropdownLists
    Set dropStd = ll.Dropdown(1)
    If dropStd.HasCheckings Then dropChecks.Push dropStd.CheckingValues

    Dim dropCust As DropdownLists
    Set dropCust = ll.Dropdown(2)
    If dropCust.HasCheckings Then dropChecks.Push dropCust.CheckingValues

    GenerationReport.FlushCheckings dropChecks

    'Build the analyses in clickGenerate
    Set anaOut = AnalysisOutput.Create(specs.AnalysisObject.Wksh(), ll)
    ' All four analysis sheets. The call used to stop after the time series
    ' tables, so the generated linelist carried no time series chart, no
    ' navigation dropdown on that sheet, and two empty sheets where the spatial
    ' and spatio-temporal analyses belong.
    anaOut.WriteAnalysis AnalysisBuildStageAll

    'Flush Phase 3: analysis checkings
    If anaOut.HasCheckings Then
        Dim anaChecks As BetterArray
        Set anaChecks = New BetterArray
        anaChecks.LowerBound = 1
        anaChecks.Push anaOut.CheckingValues
        GenerationReport.FlushCheckings anaChecks
    End If

    'Save the linelist as .xlsb with password protection
    ll.SaveLL

    'Finalise the generation report (install filter handler)
    GenerationReport.FinaliseReport

    entry.AddInfo entry.TranslateMessage("MSG_LLCreated"), "edition"

    appScope.Restore
    MsgBox entry.TranslateMessage("MSG_LLCreated"), vbInformation + vbOKOnly, PROMPT_TITLE
    Exit Sub

Cleanup:
    Dim errNumber As Long
    Dim errDesc As String
    errNumber = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    'Try to finalise whatever report was written before the error
    GenerationReport.FinaliseReport
    If Not appScope Is Nothing Then appScope.Restore
    Application.Cursor = xlDefault
    On Error GoTo 0

    If errNumber <> 0 Then
        Debug.Print "clickGenerate: "; errNumber; errDesc

        'When the linelist object exists, offer the user to view the
        'incomplete workbook or close it; otherwise show a simple error
        If Not ll Is Nothing Then
            ll.ErrorManage errDesc
        Else
            MsgBox "Generation failed: " & errDesc, _
                   vbExclamation + vbOKOnly, PROMPT_TITLE
        End If
    End If
End Sub


'@section Internal helpers
'===============================================================================

'@Description("Build a data entry worksheet from the dictionary and return the builder.")
Private Function BuildOneSheet(ByVal llshs As LLSheets, ByVal ll As Linelist, ByVal sheetName As String) As LLDataEntry
    Dim sheetType As String
    Dim layer As Byte
    Dim listBld As LLDataEntry

    sheetType = llshs.SheetInfo(sheetName)

    If sheetType = "vlist1D" Then
        layer = LLDataEntryLayerVList
    ElseIf sheetType = "hlist2D" Then
        layer = LLDataEntryLayerHList
    Else
        Exit Function
    End If

    'The builder takes the LLSheets this loop already holds. It used to build
    'its own, and so did each of the three members inside it, so one sheet cost
    'five searches of the dictionary for the same row.
    Set listBld = LLDataEntry.Create(layer, sheetName, ll, llshs)
    listBld.Build

    Set BuildOneSheet = listBld
End Function

'@Description("Extract languages from setup Translations sheet and update the setup languages dropdown.")
Private Sub ExtractAndUpdateLanguages(ByVal tradSheet As Worksheet)
    Dim setupStore As HiddenNames
    Dim langString As String
    Dim languages() As String
    Dim langValues As BetterArray
    Dim idx As Long
    Dim drop As DropdownLists

    'Read the persisted language list from the setup's Translations worksheet
    Set setupStore = HiddenNames.Create(tradSheet)

    If Not setupStore.HasName(SETUP_LANGUAGES_TAG) Then
        'Fallback: read column headers from the first ListObject on the sheet
        ExtractLanguagesFromHeaders tradSheet
        Exit Sub
    End If

    langString = setupStore.ValueAsString(SETUP_LANGUAGES_TAG)
    If LenB(langString) = 0 Then Exit Sub

    'Split semicolons-separated string into individual language names
    languages = Split(langString, ";")

    'Build BetterArray of language values (1-based)
    Set langValues = New BetterArray
    langValues.LowerBound = 1
    For idx = LBound(languages) To UBound(languages)
        If LenB(Trim$(languages(idx))) > 0 Then
            langValues.Push Trim$(languages(idx))
        End If
    Next idx

    If langValues.Length = 0 Then Exit Sub

    'Update the setup languages dropdown directly
    Set drop = DropdownLists.Create(ThisWorkbook.Worksheets(SHEET_DROPDOWNS))
    drop.Update langValues, DROP_SETUP_LANGUAGES

    'Auto-set the first language into RNG_LangSetup on the Main sheet
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_MAIN).Range("RNG_LangSetup").Value = _
        langValues.Item(langValues.LowerBound)
    On Error GoTo 0

End Sub

'@Description("Check that all required fields for generation are filled and valid.")
'@return Boolean. True when all required fields pass validation.
Private Function ValidateGenerationReadiness(ByVal entry As DesignerEntry) As Boolean
    Dim setupPath As String
    Dim llDir As String
    Dim llName As String
    Dim ribbonName As String
    Dim errors As BetterArray

    Set errors = New BetterArray
    errors.LowerBound = 1

    setupPath = entry.ValueOf("setuppath")
    llDir = entry.ValueOf("lldir")
    llName = entry.ValueOf("llname")
    ribbonName = entry.ValueOf("temppath")

    'Setup file path must be set and the file must exist on disk
    If LenB(setupPath) = 0 Then
        errors.Push "Setup file path is missing."
    ElseIf LenB(Dir(setupPath)) = 0 Then
        errors.Push "Setup file not found: " & setupPath
    End If

    'Linelist output directory must be set and exist
    If LenB(llDir) = 0 Then
        errors.Push "Linelist output directory is missing."
    ElseIf LenB(Dir(llDir, vbDirectory)) = 0 Then
        errors.Push "Output directory not found: " & llDir
    End If

    'Linelist name must be set
    If LenB(llName) = 0 Then
        errors.Push "Linelist name is missing."
    End If

    'Template ribbon must exist when configured
    If LenB(ribbonName) <> 0 Then
        If LenB(Dir(ribbonName)) = 0 Then
            errors.Push "Template ribbon file is missing: " & ribbonName
        End If
    End If

    If errors.Length > 0 Then
        entry.AddInfo entry.TranslateMessage("MSG_NotReady"), "edition"
        MsgBox errors.ToString(Separator:=vbCrLf, _
                               OpeningDelimiter:=vbNullString, _
                               ClosingDelimiter:=vbNullString), _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        ValidateGenerationReadiness = False
    Else
        ValidateGenerationReadiness = True
    End If
End Function

'@Description("Fallback: extract languages from the header row of the first ListObject on the Translations sheet.")
Private Sub ExtractLanguagesFromHeaders(ByVal tradSheet As Worksheet)
    Dim lo As ListObject
    Dim langValues As BetterArray
    Dim drop As DropdownLists

    If tradSheet.ListObjects.Count = 0 Then Exit Sub

    Set lo = tradSheet.ListObjects(1)
    If lo.HeaderRowRange Is Nothing Then Exit Sub

    'Read all header values as potential languages
    Set langValues = New BetterArray
    langValues.LowerBound = 1
    langValues.FromExcelRange lo.HeaderRowRange, _
                              DetectLastRow:=False, DetectLastColumn:=False

    If langValues.Length = 0 Then Exit Sub

    'Update the setup languages dropdown directly
    Set drop = DropdownLists.Create(ThisWorkbook.Worksheets(SHEET_DROPDOWNS))
    drop.Update langValues, DROP_SETUP_LANGUAGES

    'Auto-set the first language into RNG_LangSetup on the Main sheet
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_MAIN).Range("RNG_LangSetup").Value = _
        langValues.Item(langValues.LowerBound)
    On Error GoTo 0

End Sub
