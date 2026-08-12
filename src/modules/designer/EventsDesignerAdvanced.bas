Attribute VB_Name = "EventsDesignerAdvanced"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Non-core ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, DesignerEntry, EventsDesignerCore, RibbonDev, LLGeo, ApplicationState, OSFiles, HiddenNames, BetterArray, DropdownLists, LinelistSpecs, Linelist, LLDataEntry, LLSheets, AnalysisOutput, Checking, GenerationReport, SetupTranslationsTable
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

'Non-core ribbon logics are callbacks whose absence will not fire a
'warning at workbook opening on the designer. They only execute in
'response to explicit user actions (onAction), never at ribbon load
'time (getLabel, getPressed, getVisible).
'
'The DesignerEntry every callback works through is the shared one,
'EventsDesignerCore.EntryManager(). It carries the held translator with it,
'so a press reads no translation table.

Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_MAIN As String = "Main"
Private Const SHEET_DROPDOWNS As String = "__dropdowns"
Private Const PROMPT_TITLE As String = "Designer"

Private Const SHEET_TRANSLATIONS As String = "Translations"

'Dropdown name used by DesignerPreparation for setup languages
Private Const DROP_SETUP_LANGUAGES As String = "__setup_languages"

'Leading shape of the internal tag columns of a setup translations table.
'The fallback header read drops these before the dropdown update.
Private Const INTERNAL_TAG_LEAD As String = "__"


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

    'Preparation re-imports the translation tables, so the held pair is stale.
    EventsDesignerCore.ResetDesignerCaches

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

    Set entry = EventsDesignerCore.EntryManager()
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
    Dim prep As DesignerPreparation
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

    Set entry = EventsDesignerCore.EntryManager()

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

    'A new setup file brings its own __formatter sheet, so the designer's
    'copy stops being the live one. The styles import button sets the flag
    'again when the user wants the designer's formatter for this setup.
    Set prep = DesignerPreparation.Create(ThisWorkbook)
    prep.FormatterImported = False

    'Extract languages from the setup Translations worksheet HiddenNames
    'and update the setup languages dropdown for the designer
    ExtractAndUpdateLanguages tradSheet, entry

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

    Set entry = EventsDesignerCore.EntryManager()
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

    Set entry = EventsDesignerCore.EntryManager()
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

    Set entry = EventsDesignerCore.EntryManager()
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

    Set entry = EventsDesignerCore.EntryManager()

    'Initialise the generation report on the designer __check sheet
    GenerationReport.InitReport ThisWorkbook

    'The entry checks are the report's first bundle. Every fault carries
    'the error scope, so a checking with any entry aborts the run with the
    'report sheet shown.
    Dim entryChecks As Checking
    Set entryChecks = entry.Validate()

    If entryChecks.Length > 0 Then
        Dim entryBatch As BetterArray
        Set entryBatch = New BetterArray
        entryBatch.LowerBound = 1
        entryBatch.Push entryChecks
        GenerationReport.FlushCheckings entryBatch

        entry.AddInfo entry.TranslateMessage("MSG_NotReady"), "edition"
        GenerationReport.FinaliseReport
        appScope.Restore
        MsgBox entry.TranslateMessage("MSG_NotReady"), _
               vbExclamation + vbOKOnly, PROMPT_TITLE
        Exit Sub
    End If

    setupPath = entry.ValueOf("setuppath")

    'Build the linelist: Prepare creates the output workbook and hands it to
    'InitTransfer, which fills it from the setup file and from this designer.
    entry.AddInfo entry.TranslateMessage("MSG_ReadSetup"), "edition"

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

'@Description("Read the language names of a setup Translations sheet.")
'@details
'The one shared language extraction: the Multi group reads a setup's
'languages per row through this routine too. The persisted HiddenNames
'list of the sheet wins; the fallback reads the header row of the first
'ListObject and drops the internal tag columns (__TagInternal__ is
'machinery of the setup table, and it used to land in the dropdown).
'@param tradSheet Worksheet. The Translations worksheet of a setup workbook.
'@return BetterArray. Language names (1-based). Empty when the sheet carries none.
Public Function SetupLanguages(ByVal tradSheet As Worksheet) As BetterArray
    Dim setupStore As HiddenNames
    Dim languagesTag As String
    Dim langString As String
    Dim languages() As String
    Dim langValues As BetterArray
    Dim headerValues As BetterArray
    Dim headerText As String
    Dim lo As ListObject
    Dim idx As Long

    Set langValues = New BetterArray
    langValues.LowerBound = 1
    Set SetupLanguages = langValues

    'The HiddenName key belongs to SetupTranslationsTable, the class that
    'writes the language list on the setup's Translations sheet.
    languagesTag = SetupTranslationsTable.LanguagesNameId

    'Read the persisted language list from the setup's Translations worksheet
    Set setupStore = HiddenNames.Create(tradSheet)

    If setupStore.HasName(languagesTag) Then
        langString = setupStore.ValueAsString(languagesTag)
        If LenB(langString) > 0 Then
            'Split the semicolon-separated string into language names
            languages = Split(langString, ";")
            For idx = LBound(languages) To UBound(languages)
                If LenB(Trim$(languages(idx))) > 0 Then
                    langValues.Push Trim$(languages(idx))
                End If
            Next idx
            If langValues.Length > 0 Then Exit Function
        End If
    End If

    'Fallback: read the header row of the first ListObject on the sheet
    If tradSheet.ListObjects.Count = 0 Then Exit Function

    Set lo = tradSheet.ListObjects(1)
    If lo.HeaderRowRange Is Nothing Then Exit Function

    Set headerValues = New BetterArray
    headerValues.LowerBound = 1
    headerValues.FromExcelRange lo.HeaderRowRange, _
                                DetectLastRow:=False, DetectLastColumn:=False

    'The languages are the header row minus the internal tag columns
    For idx = headerValues.LowerBound To headerValues.UpperBound
        headerText = Trim$(CStr(headerValues.Item(idx)))
        If LenB(headerText) > 0 Then
            If Left$(headerText, Len(INTERNAL_TAG_LEAD)) <> INTERNAL_TAG_LEAD Then
                langValues.Push headerText
            End If
        End If
    Next idx
End Function

'@Description("Update the setup languages dropdown from a setup Translations sheet and auto-select the first language.")
Private Sub ExtractAndUpdateLanguages(ByVal tradSheet As Worksheet, ByVal entry As DesignerEntry)
    Dim langValues As BetterArray
    Dim drop As DropdownLists

    Set langValues = SetupLanguages(tradSheet)
    If langValues.Length = 0 Then Exit Sub

    'Update the setup languages dropdown directly
    Set drop = DropdownLists.Create(ThisWorkbook.Worksheets(SHEET_DROPDOWNS))
    drop.Update langValues, DROP_SETUP_LANGUAGES

    'Auto-select the first setup language (owner decision). The write goes
    'through the entry so the range resolution lives in one place, and
    'Validate is the net under a value the user never touches.
    entry.AddInfo langValues.Item(langValues.LowerBound), "setuplang"
End Sub
