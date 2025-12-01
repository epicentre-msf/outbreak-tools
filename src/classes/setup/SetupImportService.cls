VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetupImportService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Improved setup import service orchestrating workbook synchronisation"

'@PredeclaredId
'@Folder("Setup")
'@ModuleDescription("Improved setup import service orchestrating workbook synchronisation")
'@IgnoreModule ProcedureNotUsed, UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation
'@depends ProjectError, ApplicationState, IApplicationState, BetterArray, CustomTable, ICustomTable, LLExport, ILLExport, LLdictionary, ILLdictionary, IPasswords, Analysis, IAnalysis, SetupTranslationsTable, ISetupTranslationsTable

Option Explicit

Implements ISetupImportService

'@section State and constants
'===============================================================================
Private Const CLASS_NAME As String = "SetupImportService"
Private Const EXPORT_SHEET_NAME As String = "Exports"
Private Const DICTIONARY_SHEET_NAME As String = "Dictionary"
Private Const ANALYSIS_SHEET_NAME As String = "Analysis"
Private Const CHOICES_SHEET_NAME As String = "Choices"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"
Private Const REGISTRY_SHEET_NAME As String = "__updated"
'Default host anchor positions for setup tables using domain objects
Private Const DICTIONARY_HOST_START_ROW As Long = 5
Private Const DICTIONARY_HOST_START_COLUMN As Long = 1
Private Const CHOICES_HOST_START_ROW As Long = 4
Private Const CHOICES_HOST_START_COLUMN As Long =1
Private Const EXPORT_HOST_START_ROW As Long = 4
Private Const EXPORT_HOST_START_COLUMN As Long = 1
Private Const TRANSLATIONS_TABLE_NAME As String = "Tab_Translations"
'Incoming workbooks default to top-left anchored data
Private Const SOURCE_START_ROW As Long = 1
Private Const SOURCE_START_COLUMN As Long = 1

Private Type TSetupImportState
    importPath As String                    'Path to the source setup workbook
    progressDisplay As Object               'UI element displaying progress messages
    importWorkbook As Workbook              'Lazy loaded reference to the imported workbook
    appScope As IApplicationState           'Application state guard ensuring Excel settings are restored
    hasCaption As Boolean
    exportPath As String
    exportWorkbook As Workbook
    displayPrompt As Boolean
    lastExportFile As String
End Type

Private this As TSetupImportState

Private Sub Class_Initialize()
    this.displayPrompt = True
    this.lastExportFile = vbNullString
End Sub


'@section Factory helpers
'===============================================================================
'@sub-title Create a ready-to-use service instance
'@param importPath String path to the source setup workbook.
'@param progressDisplay Object receiving textual updates (must expose Caption).
'@return ISetupImportService configured with the provided path and progress display.
Public Function Create(ByVal importPath As String, Optional ByVal progressDisplay As Object) As ISetupImportService
    Dim instance As SetupImportService

    ValidateProgressDisplay progressDisplay
    ValidateImportPath importPath

    Set instance = New SetupImportService
    instance.Path = Trim$(importPath)
    Set instance.ProgressObject = progressDisplay

    Set Create = instance.Self
End Function


'@section Public configuration
'===============================================================================
'@sub-title Retrieve configured import path.
Public Property Get Path() As String
    Path = this.importPath
End Property

'@sub-title Store the path to the setup workbook.
Public Property Let Path(ByVal value As String)
    ValidateImportPath value
    this.importPath = Trim$(value)
End Property

'@sub-title Retrieve the progress display object.
Public Property Get ProgressObject() As Object
    Set ProgressObject = this.progressDisplay
End Property

'@sub-title Store the progress display object reference.
Public Property Set ProgressObject(ByVal value As Object)
    ValidateProgressDisplay value
    Set this.progressDisplay = value
End Property

'@sub-title Self reference implementing the interface.
Public Property Get Self() As ISetupImportService
    Set Self = Me
End Property

'@sub-title Allow callers to toggle export folder prompts (primarily for tests).
Public Property Let DisplayPrompts(ByVal state As Boolean)
    this.displayPrompt = state
End Property

Public Property Get DisplayPrompts() As Boolean
    DisplayPrompts = this.displayPrompt
End Property

'@sub-title Configure the export folder when bypassing UI prompts.
Public Sub SetExportFolder(ByVal folderPath As String)
    this.exportPath = Trim$(folderPath)
    If LenB(this.exportPath) = 0 Then
        this.exportPath = vbNullString
    End If
End Sub

'@sub-title Retrieve the last export file saved through this service.
Public Property Get LastExportFile() As String
    LastExportFile = this.lastExportFile
End Property


'@section Application state coordination
'===============================================================================
'@sub-title Guard Excel with the reusable ApplicationState scope.
Private Function ApplicationScope() As IApplicationState
    If this.appScope Is Nothing Then
        Set this.appScope = ApplicationState.Create(Application)
    End If

    Set ApplicationScope = this.appScope
End Function

'@sub-title Apply the busy state when heavy work starts.
Private Sub EnterBusyState()
    Dim scope As IApplicationState

    Set scope = ApplicationScope()
    If Not scope.IsBusy Then
        scope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, blockSecurity:=True
    End If
End Sub

'@sub-title Restore Excel configuration when work completes.
Private Sub LeaveBusyState(Optional ByVal silent As Boolean = False)
    If this.appScope Is Nothing Then Exit Sub
    
    If Not this.appScope.IsBusy Then
        Set this.appScope = Nothing
        Exit Sub
    End If

    On Error GoTo RestoreFailed
        this.appScope.Restore
    On Error GoTo 0
    Set this.appScope = Nothing
    Exit Sub

RestoreFailed:
    Set this.appScope = Nothing
    If Not silent Then
        Err.Raise Err.Number, CLASS_NAME, "Failed to restore application state: " & Err.Description
    Else
        Err.Clear
    End If
End Sub


'@section Core workflow
'===============================================================================
'@sub-title Validate the incoming request prior to import.
Public Sub Check(ByVal importDictionary As Boolean, _
                 ByVal importChoices As Boolean, _
                 ByVal importExports As Boolean, _
                 ByVal importAnalysis As Boolean, _
                 ByVal importTranslations As Boolean, _
                 Optional ByVal cleanSetup As Boolean = False)

    Dim sourcePath As String
    Dim fileProbe As String
    Dim opened As Workbook
    Dim failureDetails As String

    EnsureImportSelection importDictionary, importChoices, importExports, importAnalysis, importTranslations, cleanSetup

    If cleanSetup Then Exit Sub

    sourcePath = Path
    EnsureImportPathConfigured sourcePath
    EnsureImportFileExists sourcePath

    EnterBusyState
    On Error GoTo CheckFailed

        On Error Resume Next
            Set opened = Workbooks.Open(sourcePath)
            failureDetails = Err.Description
            Err.Clear
        On Error GoTo CheckFailed

        If opened Is Nothing Then
            ReportAndThrow ProjectError.SomethingWentWrong, _
                            BuildOpenFailureMessage(sourcePath, failureDetails)
        End If

    CheckCleanup:
        CloseWorkbook opened
        LeaveBusyState
        Exit Sub

CheckFailed:
    failureDetails = Err.Description
    CloseWorkbook opened
    LeaveBusyState True
    Err.Raise Err.Number, Err.Source, failureDetails
End Sub

'@sub-title Clean target worksheets ahead of import.
Public Sub Clean(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
    Dim workbook As Workbook
    Dim idx As Long
    Dim sheetName As String
    Dim targetSheet As Worksheet
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set workbook = ThisWorkbook
    ValidatePasswords pass
    Set sheetsList = EnsureSheetsList(sheetsList)

    EnterBusyState
    On Error GoTo CleanFailed

    For idx = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = CStr(sheetsList.Item(idx))
        Set targetSheet = workbook.Worksheets(sheetName)

        pass.UnProtect sheetName
        sheetUnlocked = True

        CleanWorksheetTables targetSheet
        ClearWorksheetComments targetSheet

        ProtectWorksheet pass, sheetName
        sheetUnlocked = False
    Next idx

CleanCleanup:
    LeaveBusyState
    Exit Sub

CleanFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        If sheetUnlocked Then
            ProtectWorksheet pass, sheetName
            sheetUnlocked = False
        End If
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@sub-title Execute the import workflow.
Public Sub Import(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim sheetName As String
    Dim idx As Long
    Dim importSheet As Worksheet
    Dim hostSheet As Worksheet
    Dim importList As ListObject
    Dim hostList As ListObject
    Dim importTable As ICustomTable
    Dim hostTable As ICustomTable
    Dim sheetUnlocked As Boolean
    Dim lastSheetName As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim formatList As BetterArray

    ValidatePasswords pass
    Set sheetsList = EnsureSheetsList(sheetsList)

    EnterBusyState
    On Error GoTo ImportFailed

        PrepareImport pass, sheetsList, renameChoices:=True
        Set importBook = EnsureImportWorkbook()
        Set hostWorkbook = ThisWorkbook
        Set formatList = New BetterArray
        

        For idx = sheetsList.LowerBound To sheetsList.UpperBound
            sheetName = CStr(sheetsList.Item(idx))
            Set formatList = ResolveDictionaryFormat(sheetName)

            If WorksheetExists(importBook, sheetName) And WorksheetExists(hostWorkbook, sheetName) Then
                pass.UnProtect sheetName
                sheetUnlocked = True
                lastSheetName = sheetName

                Set importSheet = importBook.Worksheets(sheetName)
                Set hostSheet = hostWorkbook.Worksheets(sheetName)

                For Each importList In importSheet.ListObjects
                    If ListObjectExists(hostSheet, importList.Name) Then
                        Set hostList = hostSheet.ListObjects(importList.Name)
                        Set importTable = CustomTable.Create(importList)
                        Set hostTable = CustomTable.Create(hostList)
                        hostTable.Import importTable, _ 
                         keepSourceHeaders:=(sheetName = TRANSLATIONS_SHEET_NAME), _ 
                            formatHeaders:=formatList, _ 
                            insertShift:=(sheetName = ANALYSIS_SHEET_NAME), _ 
                            strictColumnSearch:=True
                    End If
                Next importList

                ProtectWorksheet pass, sheetName
                sheetUnlocked = False
            End If
        Next idx

ImportCleanup:
    PostImport pass, sheetsList
    CloseImportWorkbook
    LeaveBusyState
    Exit Sub

ImportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        If sheetUnlocked Then
            ProtectWorksheet pass, lastSheetName
            sheetUnlocked = False
        End If
        PostImport pass, sheetsList
        CloseImportWorkbook
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub


'@section Workbook-driven import
'===============================================================================
'@sub-title Execute the workbook-driven import relying on domain classes.
Public Sub ImportFromWorkbook(ByVal pass As IPasswords, Optional ByVal sheetsList As BetterArray = Nothing)
    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim resolvedSheets As BetterArray
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    ValidatePasswords pass
    'Here the exports coming from export for migration are on row 1, column 1
    PrepareImport pass, sheetsList, True, SOURCE_START_ROW, SOURCE_START_COLUMN
    Set resolvedSheets = ResolveImportSheets(sheetsList)

    EnterBusyState

    On Error GoTo ImportWorkbookFailed

        Set importBook = EnsureImportWorkbook()
        Set hostWorkbook = ThisWorkbook

        If ShouldImportSheet(resolvedSheets, DICTIONARY_SHEET_NAME) Then
            ImportDictionaryUsingClass pass, hostWorkbook, importBook
        End If

        If ShouldImportSheet(resolvedSheets, EXPORT_SHEET_NAME) Then
            ImportExportUsingClass pass, hostWorkbook, importBook
        End If

        If ShouldImportSheet(resolvedSheets, CHOICES_SHEET_NAME) Then
            ImportChoicesUsingClass pass, hostWorkbook, importBook
        End If

        If ShouldImportSheet(resolvedSheets, ANALYSIS_SHEET_NAME) Then
            ImportAnalysisUsingClass pass, hostWorkbook, importBook
        End If

        If ShouldImportSheet(resolvedSheets, TRANSLATIONS_SHEET_NAME) Then
            ImportTranslationsUsingClass pass, hostWorkbook, importBook
        End If

ImportWorkbookCleanup:
        CloseImportWorkbook
        LeaveBusyState
        Exit Sub

ImportWorkbookFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        LeaveBusyState True
        On Error Resume Next
            CloseImportWorkbook
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub

'@sub-title Import dictionary worksheet using LLdictionary.
Private Sub ImportDictionaryUsingClass(ByVal pass As IPasswords, _
                                       ByVal hostWorkbook As Workbook, _
                                       ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim dictionaryManager As ILLdictionary
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If Not WorksheetExists(hostWorkbook, DICTIONARY_SHEET_NAME) Then Exit Sub
    If Not WorksheetExists(importWorkbook, DICTIONARY_SHEET_NAME) Then Exit Sub

    Set hostSheet = hostWorkbook.Worksheets(DICTIONARY_SHEET_NAME)
    Set sourceSheet = importWorkbook.Worksheets(DICTIONARY_SHEET_NAME)

    On Error GoTo ImportFailed
        pass.UnProtect DICTIONARY_SHEET_NAME
        sheetUnlocked = True

        Set dictionaryManager = LLdictionary.Create(hostSheet, DICTIONARY_HOST_START_ROW, DICTIONARY_HOST_START_COLUMN)
        dictionaryManager.Import sourceSheet, SOURCE_START_ROW, SOURCE_START_COLUMN

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, DICTIONARY_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub

'@sub-title Import Choices using ILLChoices
Private Sub ImportChoicesUsingClass(ByVal pass As IPasswords, ByVal hostworkbook As Workbook, ByVal importworkbook As Workbook)

    Dim hostsheet As worksheet
    Dim sourcesheet As Worksheet
    Dim choicemanager As ILLChoices
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If Not WorksheetExists(hostWorkbook, CHOICES_SHEET_NAME) Then Exit Sub
    If Not WorksheetExists(importWorkbook, CHOICES_SHEET_NAME) Then Exit Sub

    Set hostSheet = hostWorkbook.Worksheets(CHOICES_SHEET_NAME)
    Set sourceSheet = importWorkbook.Worksheets(CHOICES_SHEET_NAME)

    On Error GoTo ImportFailed
    pass.UnProtect CHOICES_SHEET_NAME
    sheetUnlocked = True
    Set choicemanager = LLChoices.Create(hostSheet, CHOICES_HOST_START_ROW, CHOICES_HOST_START_COLUMN)
    choicemanager.Import sourceSheet, SOURCE_START_ROW, SOURCE_START_COLUMN

ImportCleanup:
        On Error Resume Next
        If sheetUnlocked Then
            ProtectWorksheet pass, CHOICES_SHEET_NAME
            sheetUnlocked = False
        End If
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

ImportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ImportCleanup
End Sub

'@sub-title Import export specifications via LLExport.
Private Sub ImportExportUsingClass(ByVal pass As IPasswords, _
                                   ByVal hostWorkbook As Workbook, _
                                   ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim exportsManager As ILLExport
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If Not WorksheetExists(hostWorkbook, EXPORT_SHEET_NAME) Then Exit Sub
    If Not WorksheetExists(importWorkbook, EXPORT_SHEET_NAME) Then Exit Sub

    Set hostSheet = hostWorkbook.Worksheets(EXPORT_SHEET_NAME)
    Set sourceSheet = importWorkbook.Worksheets(EXPORT_SHEET_NAME)

    On Error GoTo ImportFailed
        pass.UnProtect EXPORT_SHEET_NAME
        sheetUnlocked = True

        Set exportsManager = LLExport.Create(hostSheet, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN)
        exportsManager.ImportSpecs sourceSheet, SOURCE_START_ROW, SOURCE_START_COLUMN

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, EXPORT_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub



'@sub-title Import analysis worksheet content via Analysis class.
Private Sub ImportAnalysisUsingClass(ByVal pass As IPasswords, _
                                     ByVal hostWorkbook As Workbook, _
                                     ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim analysisManager As IAnalysis
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If Not WorksheetExists(hostWorkbook, ANALYSIS_SHEET_NAME) Then Exit Sub
    If Not WorksheetExists(importWorkbook, ANALYSIS_SHEET_NAME) Then Exit Sub

    Set hostSheet = hostWorkbook.Worksheets(ANALYSIS_SHEET_NAME)
    Set sourceSheet = importWorkbook.Worksheets(ANALYSIS_SHEET_NAME)

    On Error GoTo ImportFailed
        pass.UnProtect ANALYSIS_SHEET_NAME
        sheetUnlocked = True

        Set analysisManager = Analysis.Create(hostSheet)
        analysisManager.Import sourceSheet

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, ANALYSIS_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub

'@sub-title Import translations worksheet using CustomTable and SetupTranslationsTable.
Private Sub ImportTranslationsUsingClass(ByVal pass As IPasswords, _
                                         ByVal hostWorkbook As Workbook, _
                                         ByVal importWorkbook As Workbook)

    Dim hostSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim hostTable As ListObject
    Dim hostTableAdapter As ICustomTable
    Dim sourceTableAdapter As IDataSheet
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If Not WorksheetExists(hostWorkbook, TRANSLATIONS_SHEET_NAME) Then Exit Sub
    If Not WorksheetExists(importWorkbook, TRANSLATIONS_SHEET_NAME) Then Exit Sub

    Set hostSheet = hostWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)
    Set sourceSheet = importWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)

    Set hostTable = ResolveTranslationsList(hostSheet)

    If hostTable Is Nothing Then Exit Sub

    On Error GoTo ImportFailed
        pass.UnProtect TRANSLATIONS_SHEET_NAME
        sheetUnlocked = True

        Set hostTableAdapter = CustomTable.Create(hostTable)
        Set sourceTableAdapter = DataSheet.Create(sourceSheet, 1, 2)
        hostTableAdapter.Import sourceTableAdapter, keepSourceHeaders:=True

ImportCleanup:
        On Error Resume Next
            If sheetUnlocked Then
                ProtectWorksheet pass, TRANSLATIONS_SHEET_NAME
                sheetUnlocked = False
            End If
        On Error GoTo 0
        If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
        Exit Sub

ImportFailed:
        errNumber = Err.Number
        errSource = Err.Source
        errDescription = Err.Description
        Resume ImportCleanup
End Sub


'@sub-title Resolve the translations listobject from a worksheet.
Private Function ResolveTranslationsList(ByVal sheetRef As Worksheet, _ 
                                         Optional ByVal preferred As ListObject = Nothing) As ListObject
    Dim candidate As ListObject

    If sheetRef Is Nothing Then Exit Function

    On Error Resume Next
        If Not preferred Is Nothing Then
            Set candidate = sheetRef.ListObjects(preferred.Name)
        End If
    On Error GoTo 0

    If candidate Is Nothing Then
        On Error Resume Next
            Set candidate = sheetRef.ListObjects(TRANSLATIONS_TABLE_NAME)
        On Error GoTo 0
    End If

    If candidate Is Nothing Then
        On Error Resume Next
            If sheetRef.ListObjects.Count > 0 Then
                Set candidate = sheetRef.ListObjects(1)
            End If
        On Error GoTo 0
    End If

    Set ResolveTranslationsList = candidate
End Function

'@section Workbook driven Export
'===============================================================================
Private Sub Export()

    Dim dictionaryManager As ILLdictionary
    Dim translationManager As ISetupTranslationsTable
    Dim translationsTable As ListObject
    Dim analysisManager As IAnalysis
    Dim exportsManager As ILLExport
    Dim choicesManager As ILLChoices
    Dim exportworkbook As Workbook
    Dim hostWorkbook As Workbook
    Dim hostSheet As Worksheet
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim store As IHiddenNames


    this.lastExportFile = vbNullString

    Set exportworkbook = EnsureExportWorkbook()
    If exportworkbook Is Nothing Then Exit Sub
    Set hostWorkbook = ThisWorkbook
    If hostworkbook Is Nothing Then Exit Sub
    
    EnterBusyState
    On Error GoTo ExportFailed

    'Export the dictionary
    If WorksheetExists(hostWorkbook, DICTIONARY_SHEET_NAME) Then
        Set hostSheet = hostWorkbook.Worksheets(DICTIONARY_SHEET_NAME)
        Set dictionaryManager = LLdictionary.Create(hostSheet, DICTIONARY_HOST_START_ROW, DICTIONARY_HOST_START_COLUMN)
        dictionaryManager.Export exportworkbook, Hide:=xlSheetVisible
    End If

    'Export the choices
    If WorksheetExists(hostWorkbook, CHOICES_SHEET_NAME) Then
        Set hostSheet = hostWorkbook.Worksheets(CHOICES_SHEET_NAME)
        Set choicesManager = LLChoices.Create(hostSheet, CHOICES_HOST_START_ROW, CHOICES_HOST_START_COLUMN)
        choicesManager.Export exportworkbook, Hide:=xlSheetVisible
    End If

    'Export the exports
    If WorksheetExists(hostWorkbook, EXPORT_SHEET_NAME) Then
        Set hostSheet = hostWorkbook.Worksheets(EXPORT_SHEET_NAME)
        Set exportsManager = LLExport.Create(hostSheet, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN)
        exportsManager.ExportSpecs exportworkbook, Hide:=xlSheetVisible
    End If


    'Export the analysis
    If WorksheetExists(hostWorkbook, ANALYSIS_SHEET_NAME) Then
        Set hostSheet = hostWorkbook.Worksheets(ANALYSIS_SHEET_NAME)
        Set analysisManager = Analysis.Create(hostSheet)
        analysisManager.Export exportworkbook, Hide:=xlSheetVisible
    End If

    'Export the translations
    If WorksheetExists(hostWorkbook, TRANSLATIONS_SHEET_NAME) Then
        Set hostSheet = hostWorkbook.Worksheets(TRANSLATIONS_SHEET_NAME)
        Set translationsTable = ResolveTranslationsList(hostSheet)

        If Not translationsTable Is Nothing Then
            Set translationManager = SetupTranslationsTable.Create(translationsTable)
            translationManager.Export exportworkbook, Hide:=xlSheetVisible
        End If
    End If

    'Export hidden names in the workbook
    Set store = HiddenNames.Create(hostWorkbook)
    store.ExportNamesToWorkbook exportworkbook
    

    On Error Resume Next
    exportWorkbook.Worksheets(1).Delete
    On Error GoTo 0

ExportCleanUp:
    EnsureSavedExport
    CloseExportWorkbook
    LeaveBusyState
    Exit Sub

ExportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
        CloseExportWorkbook
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub


'@section Preparation helpers
'===============================================================================
'@sub-title Ensure the exports and dictionary sheets are aligned prior to import.
Private Sub PrepareImport(ByVal pass As IPasswords, _
                          Optional ByVal sheetsList As BetterArray, _ 
                          Optional ByVal forceSync As Boolean = True, _
                          Optional ByVal impStartRow As Long = EXPORT_HOST_START_ROW, _
                          Optional ByVal impStartColumn As Long = EXPORT_HOST_START_COLUMN, _ 
                          Optional ByVal renameChoices As Boolean = False)

    Dim importBook As Workbook
    Dim hostWorkbook As Workbook
    Dim currentExportsSheet As Worksheet
    Dim importExportsSheet As Worksheet
    Dim currentExportObj As ILLExport
    Dim importExportObj As ILLExport
    Dim currentDictionary As ILLdictionary
    Dim previousExports As Long
    Dim updatedExports As Long
    Dim counter As Long
    Dim dictUnlocked As Boolean
    Dim exportsUnlocked As Boolean
    Dim choiUnlocked As Boolean
    Dim currentChoiceData As ICustomTable
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If sheetsList Is Nothing Then Exit Sub


    If (Not sheetsList.Includes(EXPORT_SHEET_NAME)) And (Not forceSync) And (Not sheetsList.Includes(CHOICES_SHEET_NAME)) Then Exit Sub


    Set importBook = EnsureImportWorkbook()
    Set hostWorkbook = ThisWorkbook

    On Error GoTo PrepareFailed

    If sheetsList.Includes(EXPORT_SHEET_NAME) Or forceSync Then
        If WorksheetExists(importBook, EXPORT_SHEET_NAME) And WorksheetExists(importBook, DICTIONARY_SHEET_NAME) Then
            Set currentExportsSheet = hostWorkbook.Worksheets(EXPORT_SHEET_NAME)
            Set importExportsSheet = importBook.Worksheets(EXPORT_SHEET_NAME)
            Set currentExportObj = LLExport.Create(currentExportsSheet, EXPORT_HOST_START_ROW, EXPORT_HOST_START_COLUMN)
            Set importExportObj = LLExport.Create(importExportsSheet, impStartRow, impStartColumn)

            If currentExportObj.NumberOfExports < importExportObj.NumberOfExports Then
                pass.UnProtect DICTIONARY_SHEET_NAME
                dictUnlocked = True
                pass.UnProtect EXPORT_SHEET_NAME
                exportsUnlocked = True

                Set currentDictionary = LLdictionary.Create(hostWorkbook.Worksheets(DICTIONARY_SHEET_NAME), DICTIONARY_HOST_START_ROW, DICTIONARY_HOST_START_COLUMN)
                updatedExports = currentExportObj.NumberOfExports

                Do While  updatedExports < importExportObj.NumberOfExports
                    previousExports = currentExportObj.NumberOfExports
                    currentExportObj.AddRows currentDictionary
                    updatedExports = currentExportObj.NumberOfExports
                Loop
            End If
        End If
    End If

    If sheetsList.Includes(CHOICES_SHEET_NAME) And renameChoices Then
        If WorksheetExists(importBook, CHOICES_SHEET_NAME) And sheetsList.Includes(CHOICES_SHEET_NAME) Then
            pass.UnProtect CHOICES_SHEET_NAME
            choiUnlocked = True
            Set currentChoiceData = CustomTable.Create(hostworkbook.Worksheets(CHOICES_SHEET_NAME).ListObjects(1))
            currentChoiceData.RenameColumn "Label", "Formula Label", matchCase:=False
            currentChoiceData.RenameColumn "Translated Label", "Label", matchCase:=False
        End If
    End If

PrepareCleanup:
    On Error Resume Next
        If dictUnlocked Then
            pass.Protect DICTIONARY_SHEET_NAME
            dictUnlocked = False
        End If
        If exportsUnlocked Then
            pass.Protect EXPORT_SHEET_NAME
            exportsUnlocked = False
        End If
        If choiUnlocked Then
            pass.Protect CHOICES_SHEET_NAME
            choiUnlocked = False
        End If

    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

PrepareFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume PrepareCleanup
End Sub


Private Sub PostImport(ByVal pass As IPasswords, _
                        Optional ByVal sheetsList As BetterArray)

    If sheetsList Is Nothing Then Exit Sub
    If Not sheetsList.Includes(CHOICES_SHEET_NAME) Then Exit Sub

    Dim importBook As Workbook
    Dim hostworkbook As Workbook
    Dim choiUnlocked As Boolean
    Dim currentChoiceData As ICustomTable
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo PostFailed

    Set importBook = EnsureImportWorkbook()
    Set hostworkbook = ThisWorkbook

    If WorksheetExists(importBook, CHOICES_SHEET_NAME) Then
        pass.UnProtect CHOICES_SHEET_NAME
        choiUnlocked = True
        Set currentChoiceData = CustomTable.Create(hostworkbook.Worksheets(CHOICES_SHEET_NAME).ListObjects(1))
        currentChoiceData.RenameColumn "Label", "Translated Label", matchCase:=False
        currentChoiceData.RenameColumn "Formula Label", "Label", matchCase:=False
    End If

PostCleanUp:
    On Error Resume Next
        If choiUnlocked Then
            pass.Protect CHOICES_SHEET_NAME
            choiUnlocked = False
        End If
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

PostFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume PostCleanUp
End Sub

'@section Validation helpers
'===============================================================================
'@sub-title Resolve the requested sheets set, defaulting to core setup worksheets.
Private Function ResolveImportSheets(ByVal sheetsList As BetterArray) As BetterArray
    If sheetsList Is Nothing Then
        Set ResolveImportSheets = DefaultImportSheets()
    Else
        Set ResolveImportSheets = EnsureSheetsList(sheetsList)
    End If
End Function

'@sub-title Default sheet list covering the core setup worksheets.
Private Function DefaultImportSheets() As BetterArray
    Dim defaults As BetterArray

    Set defaults = New BetterArray
    defaults.LowerBound = 1
    defaults.Push DICTIONARY_SHEET_NAME, CHOICES_SHEET_NAME, EXPORT_SHEET_NAME, ANALYSIS_SHEET_NAME, TRANSLATIONS_SHEET_NAME

    Set DefaultImportSheets = defaults
End Function

'@sub-title Determine whether a sheet should participate in the import.
Private Function ShouldImportSheet(ByVal sheetsList As BetterArray, ByVal sheetName As String) As Boolean
    Dim idx As Long
    Dim candidate As String

    If sheetsList Is Nothing Then
        ShouldImportSheet = True
        Exit Function
    End If

    For idx = sheetsList.LowerBound To sheetsList.UpperBound
        candidate = CStr(sheetsList.Item(idx))
        If StrComp(candidate, sheetName, vbTextCompare) = 0 Then
            ShouldImportSheet = True
            Exit Function
        End If
    Next idx
End Function

'@sub-title Ensure at least one import option is selected.
Private Sub EnsureImportSelection(ByVal importDictionary As Boolean, _
                                  ByVal importChoices As Boolean, _
                                  ByVal importExports As Boolean, _
                                  ByVal importAnalysis As Boolean, _
                                  ByVal importTranslations As Boolean, _
                                  ByVal cleanSetup As Boolean)

    Const OPTIONS_WARNING As String = "Please select at least one import option (Dictionary, Choices, Exports, Analysis or Translations)."

    If Not importDictionary And Not importChoices And Not importExports _
       And Not importAnalysis And Not importTranslations Then
        If cleanSetup Then
            ReportAndThrow ProjectError.InvalidArgument, _
                            "Clean operation requires at least one worksheet to process."
        Else
            ReportAndThrow ProjectError.InvalidArgument, OPTIONS_WARNING
        End If
    End If
End Sub

'@sub-title Ensure the import path has been configured.
Private Sub EnsureImportPathConfigured(ByVal value As String)
    If LenB(Trim$(value)) = 0 Then
        ReportAndThrow ProjectError.InvalidArgument, "No setup file path is configured for this import."
    End If
End Sub

'@sub-title Validate availability of requested passwords object.
Private Sub ValidatePasswords(ByVal pass As IPasswords)
    If pass Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Passwords handler is required for worksheet protection."
    End If
End Sub

'@sub-title Ensure the sheets list is initialised.
Private Function EnsureSheetsList(ByVal sheetsList As BetterArray) As BetterArray
    If sheetsList Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Sheets list cannot be Nothing."
    End If

    If sheetsList.Length = 0 Then
        ThrowError ProjectError.InvalidArgument, "Sheets list cannot be empty."
    End If

    Set EnsureSheetsList = sheetsList
End Function

'@sub-title Confirm the import path string is valid.
Private Sub ValidateImportPath(ByVal pathValue As String)
    If LenB(Trim$(pathValue)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Import path cannot be blank."
    End If
End Sub

'@sub-title Ensure the progress display object exposes a Caption property.
Private Sub ValidateProgressDisplay(ByVal progressDisplay As Object)

    If (Not progressDisplay Is Nothing) Then
        If TestProperty(progressDisplay, this.hasCaption, "Caption") Then Exit Sub
        If TestProperty(progressDisplay, this.hasCaption, "Value") Then Exit Sub
    End If
Exit Sub

InvalidDisplay:
    Err.Clear
    ThrowError ProjectError.InvalidArgument, "Progress display object must expose a Caption or Value property."
End Sub

Private Function TestProperty(ByVal Obj As Object, ByRef hasProp As Boolean, Optional ByVal propName As String = "Caption") As Boolean

    On Error Resume Next
        If propName = "Caption" Then
            Obj.Caption = ""
        ElseIf propName = "Value" Then
            Obj.Value = ""
        End If

        hasProp = (Err.Number = 0)
        TestProperty = hasProp
    On Error GoTo 0
End Function

'@sub-title Ensure the import workbook file exists before opening.
Private Sub EnsureImportFileExists(ByVal filePath As String)
    Dim resolved As String

    On Error Resume Next
        resolved = Dir(filePath)
    On Error GoTo 0

    If LenB(resolved) = 0 Then
        ReportAndThrow ProjectError.ElementNotFound, _
                        "No setup file was found at '" & filePath & "'."
    End If
End Sub


'@section Workbook helpers
'===============================================================================
'@sub-title Obtain the imported workbook, opening it when needed.
Private Function EnsureImportWorkbook() As Workbook
    Dim failureDetails As String

    EnsureImportPathConfigured this.importPath

    If this.importWorkbook Is Nothing Then
        On Error Resume Next
            Set this.importWorkbook = Workbooks.Open(this.importPath)
            failureDetails = Err.Description
            Err.Clear
        On Error GoTo 0

        If this.importWorkbook Is Nothing Then
            ReportAndThrow ProjectError.SomethingWentWrong, _
                            BuildOpenFailureMessage(this.importPath, failureDetails)
        End If
    End If

    Set EnsureImportWorkbook = this.importWorkbook
End Function

Private Function EnsureExportWorkbook() As Workbook

    Dim io As IOSFiles

    If (this.exportWorkbook Is Nothing) Then
        If LenB(this.exportPath) = 0 Then
            If this.displayPrompt Then
                Set io = OSFiles.Create()

                io.LoadFolder

                If io.HasValidFolder() Then
                    this.exportPath = io.Folder()
                Else
                    Set EnsureExportWorkbook = Nothing
                    Exit Function
                End If
            Else
                Set EnsureExportWorkbook = Nothing
                Exit Function
            End If
        End If

        If LenB(this.exportPath) = 0 Then
            Set EnsureExportWorkbook = Nothing
            Exit Function
        End If

        Set this.exportWorkbook = Workbooks.Add
        ActiveWindow.WindowState = xlMinimized
    End If

    Set EnsureExportWorkbook = this.exportWorkbook
End Function


Private Sub EnsureSavedExport()

    Dim exportworkbook As Workbook
    Dim filePath As String

    Set exportworkbook = this.exportWorkbook
    If exportworkbook Is Nothing Then Exit Sub
    If LenB(this.exportPath) = 0 Then Exit Sub

    filePath = BuildExportFilePath()

    exportworkbook.SaveAs filePath, _
    CreateBackup:=False, _
    ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    this.lastExportFile = filePath
End Sub

Private Function ResolveDictionaryFormat(ByVal sheetName As String) As BetterArray

    Dim lst As BetterArray
    Set lst = New BetterArray
    If sheetName = DICTIONARY_SHEET_NAME Then
        lst.Push "Formatting Condition", _ 
                "Formatting Values", _ 
                "Variable Name", _ 
                "Control", "Main Label", "Lock Cells", "Dev Comments"
    End If

    Set ResolveDictionaryFormat = lst.Clone()

End Function

Private Function BuildExportFilePath() As String
    Dim filePath As String

    filePath = Replace(ThisWorkbook.Name, ".xlsb", vbNullString) & "_export_" & Format(Now(), "yyyymmdd") & ".xlsx"
    BuildExportFilePath = this.exportPath & Application.PathSeparator & filePath
End Function

'@sub-title Close and release the imported workbook.
Private Sub CloseImportWorkbook()
    CloseWorkbook this.importWorkbook
    Set this.importWorkbook = Nothing
End Sub

Private Sub CloseExportWorkbook()
    CloseWorkbook this.exportworkbook
    Set this.exportworkbook = Nothing
    this.exportPath = vbNullString
End Sub


'@sub-title Close workbook safely without raising errors.
Private Sub CloseWorkbook(ByVal workbookRef As Workbook)
    On Error Resume Next
        If Not workbookRef Is Nothing Then
            workbookRef.Close savechanges:=False
        End If
    On Error GoTo 0
End Sub

'@sub-title Detect whether a worksheet exists within a workbook.
Private Function WorksheetExists(ByVal workbookRef As Workbook, ByVal sheetName As String) As Boolean
    Dim sheetRef As Worksheet

    If workbookRef Is Nothing Then Exit Function

    On Error Resume Next
        Set sheetRef = workbookRef.Worksheets(sheetName)
    On Error GoTo 0

    WorksheetExists = Not (sheetRef Is Nothing)
End Function

'@sub-title Resolve the registry worksheet if present.
Private Function ResolveRegistrySheet(ByVal workbookRef As Workbook) As Worksheet
    On Error Resume Next
        Set ResolveRegistrySheet = workbookRef.Worksheets(REGISTRY_SHEET_NAME)
    On Error GoTo 0
End Function

'@sub-title Detect whether a list object exists within a worksheet.
Private Function ListObjectExists(ByVal sheetRef As Worksheet, ByVal listName As String) As Boolean
    Dim listRef As ListObject

    If sheetRef Is Nothing Then Exit Function

    On Error Resume Next
        Set listRef = sheetRef.ListObjects(listName)
    On Error GoTo 0

    ListObjectExists = Not (listRef Is Nothing)
End Function


'@section Worksheet helpers
'===============================================================================
'@sub-title Clean all list objects within a worksheet.
Private Sub CleanWorksheetTables(ByVal targetSheet As Worksheet)
    Dim listRef As ListObject
    Dim table As ICustomTable

    For Each listRef In targetSheet.ListObjects
        Set table = CustomTable.Create(listRef)
        table.Clean
    Next listRef
End Sub

'@sub-title Remove both classic and threaded comments from a worksheet.
Private Sub ClearWorksheetComments(ByVal targetSheet As Worksheet)
    Dim usedCells As Range

    On Error Resume Next
        Set usedCells = targetSheet.UsedRange
    On Error GoTo 0

    If usedCells Is Nothing Then Exit Sub

    On Error Resume Next
        usedCells.ClearComments
        Err.Clear
    On Error GoTo 0
End Sub

'@sub-title Protect worksheet with default rules.
Private Sub ProtectWorksheet(ByVal pass As IPasswords, ByVal sheetName As String)
    pass.Protect sheetName, _
                      (sheetName = TRANSLATIONS_SHEET_NAME Or sheetName = ANALYSIS_SHEET_NAME), _
                      (sheetName = TRANSLATIONS_SHEET_NAME)
End Sub


'@section Messaging helpers
'===============================================================================
'@sub-title Display informational text through the progress object.
Private Sub WriteInfo(ByVal message As String)
    If Not this.hasCaption Then Exit Sub
    If this.progressDisplay Is Nothing Then Exit Sub

    On Error Resume Next
        this.progressDisplay.Caption = message
        this.progressDisplay.Value = message
    On Error GoTo 0
End Sub

'@sub-title Update a textual progress bar representation.
Private Sub WriteProgress(ByVal percentage As Integer)
    Const TOTAL_BARS As Integer = 30
    Dim completedBars As Integer
    Dim clamped As Integer

    If this.progressDisplay Is Nothing Then Exit Sub
    clamped = Application.WorksheetFunction.Min(percentage, 100)
    completedBars = TOTAL_BARS * Round(clamped / 100, 1)

    On Error Resume Next
        this.progressDisplay.Caption = "[" & String$(completedBars, "|") & Space$(TOTAL_BARS - completedBars) & "] " & CStr(clamped) & "%"
    On Error GoTo 0
End Sub

'@sub-title Combine failure details when opening a workbook fails.
Private Function BuildOpenFailureMessage(ByVal sourcePath As String, ByVal failureDetails As String) As String
    Dim message As String

    message = "Unable to open the setup workbook at '" & sourcePath & "'."
    If LenB(Trim$(failureDetails)) > 0 Then
        message = message & " (" & failureDetails & ")"
    End If

    BuildOpenFailureMessage = message
End Function


'@section Error handling
'===============================================================================
'@sub-title Report message then raise a project error.
Private Sub ReportAndThrow(ByVal errNumber As ProjectError, ByVal message As String)
    WriteInfo message
    ThrowError errNumber, message
End Sub

'@sub-title Raise a ProjectError-based exception.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub

'@sub-title Cleanup resources on destruction.
Private Sub Class_Terminate()
    On Error Resume Next
        CloseImportWorkbook
        LeaveBusyState True
    On Error GoTo 0
    Set this.progressDisplay = Nothing
    Set this.appScope = Nothing
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ISetupImportService_Check(ByVal importDictionary As Boolean, _
                                      ByVal importChoices As Boolean, _
                                      ByVal importExports As Boolean, _
                                      ByVal importAnalysis As Boolean, _
                                      ByVal importTranslations As Boolean, _
                                      Optional ByVal cleanSetup As Boolean = False)

    Check importDictionary, importChoices, importExports, importAnalysis, importTranslations, cleanSetup
End Sub

Private Sub ISetupImportService_Clean(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
    Clean pass, sheetsList
End Sub

Private Sub ISetupImportService_Import(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
    Import pass, sheetsList
End Sub

Private Sub ISetupImportService_ImportFromWorkbook(ByVal pass As IPasswords, Optional ByVal sheetsList As BetterArray = Nothing)
    ImportFromWorkbook pass, sheetsList
End Sub

Private Sub ISetupImportService_Export()
    Export
End Sub

Private Property Get ISetupImportService_DisplayPrompts() As Boolean
    ISetupImportService_DisplayPrompts = DisplayPrompts
End Property

Private Property Let ISetupImportService_DisplayPrompts(ByVal state As Boolean)
    DisplayPrompts = state
End Property

Private Sub ISetupImportService_SetExportFolder(ByVal folderPath As String)
    SetExportFolder folderPath
End Sub

Private Property Get ISetupImportService_LastExportFile() As String
    ISetupImportService_LastExportFile = LastExportFile
End Property

Private Property Get ISetupImportService_Path() As String
    ISetupImportService_Path = Path
End Property

Private Property Let ISetupImportService_Path(ByVal value As String)
    Path = value
End Property

Private Property Get ISetupImportService_ProgressObject() As Object
    Set ISetupImportService_ProgressObject = ProgressObject
End Property

Private Property Set ISetupImportService_ProgressObject(ByVal value As Object)
    Set ProgressObject = value
End Property

Private Property Get ISetupImportService_Self() As ISetupImportService
    Set ISetupImportService_Self = Self
End Property
