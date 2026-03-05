Attribute VB_Name = "EventsDesignerAdvanced"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Non-core ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, IDesignerPreparation, DesignerEntry, IDesignerEntry, RibbonDev, LLGeo, ILLGeo, ApplicationState, IApplicationState, OSFiles, IOSFiles, HiddenNames, IHiddenNames, BetterArray, DropdownLists, IDropdownLists, DesignerImportService, IDesignerImportService, LinelistSpecs, ILinelistSpecs, Linelist, ILinelist, ListBuilder, IListBuilder, LLSheets, ILLSheets
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
Public Sub clickDevInitialize(ByRef control As IRibbonControl)
    Dim prep As IDesignerPreparation
    Dim appScope As IApplicationState

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
        Debug.Print "clickDevInitialize: "; errNumber; errDesc
        MsgBox "Unable to initialise designer: " & errDesc, _
               vbExclamation + vbOKOnly, PROMPT_TITLE
    End If
End Sub


'@section Manage group callbacks
'===============================================================================

'@Description("Clear all geobase data from the Geo worksheet.")
'@EntryPoint
Public Sub clickDelGeo()
    Dim geoSheet As Worksheet
    Dim geo As ILLGeo
    Dim appScope As IApplicationState

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
Public Sub clickClearEnt()
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

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
    Dim io As IOSFiles
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState
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
    Dim io As IOSFiles
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

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
    Dim io As IOSFiles
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

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
    Dim io As IOSFiles
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

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
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState
    Dim importService As IDesignerImportService
    Dim specs As ILinelistSpecs
    Dim ll As ILinelist
    Dim setupPath As String
    Dim sheetLists As BetterArray
    Dim counter As Long
    Dim anaOut As IAnalysisOutput

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, busyCursor:=xlNorthWestArrow

    Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(SHEET_MAIN))

    'Run readiness checks; exit silently when requirements are not met
    If Not ValidateGenerationReadiness(entry) Then GoTo Cleanup

    setupPath = entry.ValueOf("setuppath")

    'Import all setup components from the setup file into the designer
    entry.AddInfo entry.TranslateMessage("MSG_ReadSetup"), "edition"

    Set importService = DesignerImportService.Create(ThisWorkbook)
    importService.ImportFromSetup setupPath

    'Prepare specifications: create output workbook, export, translate
    'The specs is created on this workbook. But after import/export
    'the internal values will shift from thisworkbook to the linelistworkbook.
    
    Set specs = LinelistSpecs.Create(ThisWorkbook)
    specs.Prepare importService

    'After the preparation step of the specifications, internal specifications
    'object shift focus from the designer to the linelist workbook as they 
    'are now exported.

    'Build the output linelist workbook (sheets, temp sheets, admin, code transfer)
    Set ll = Linelist.Create(specs)
    ll.Prepare

    'Build the first data entry worksheet (sections, variables, formatting)
    Set sheetLists = ll.Dictionary.UniqueValues("sheet name")

    If sheetLists.Length > 0 Then 
        For counter = sheetLists.LowerBound To sheetLists.UpperBound
            BuildOneSheet specs, ll, sheetLists.Item(counter)
        Next
    End If

    'Build the analyses in clickGenerate
    Set anaOut = AnalysisOutput.Create(specs.AnalysisObject.Wksh(), ll)
    anaOut.WriteAnalysis AnalysisScopeNormal

    'Save the linelist as .xlsb with password protection
    ll.SaveLL

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

'@Description("Build the first data entry worksheet from the dictionary.")
Private Sub BuildOneSheet(ByVal specs As ILinelistSpecs, ByVal ll As ILinelist, ByVal sheetName As String)
    Dim dict As ILLdictionary
    Dim llshs As ILLSheets
    Dim sheetType As String
    Dim layer As Byte
    Dim listBld As IListBuilder

    Set dict = specs.Dictionary
    Set llshs = LLSheets.Create(dict)

    sheetType = llshs.SheetInfo(sheetName)

    If sheetType = "vlist1D" Then
        layer = ListBuilderLayerVList
    ElseIf sheetType = "hlist2D" Then
        layer = ListBuilderLayerHList
    Else
        Exit Sub
    End If

    Set listBld = ListBuilder.Create(layer, sheetName, ll)
    listBld.Build
End Sub

'@Description("Extract languages from setup Translations sheet and update the setup languages dropdown.")
Private Sub ExtractAndUpdateLanguages(ByVal tradSheet As Worksheet)
    Dim setupStore As IHiddenNames
    Dim langString As String
    Dim languages() As String
    Dim langValues As BetterArray
    Dim idx As Long
    Dim drop As IDropdownLists

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
Private Function ValidateGenerationReadiness(ByVal entry As IDesignerEntry) As Boolean
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
    Dim drop As IDropdownLists

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
