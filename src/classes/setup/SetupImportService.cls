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
'@depends ProjectError, ApplicationState, IApplicationState, BetterArray, CustomTable, ICustomTable, LLExport, ILLExport, LLdictionary, ILLdictionary, IPasswords

Option Explicit

Implements ISetupImportService

'@section State and constants
'===============================================================================
Private Const CLASS_NAME As String = "SetupImportService"
Private Const EXPORT_SHEET_NAME As String = "Exports"
Private Const DICTIONARY_SHEET_NAME As String = "Dictionary"
Private Const ANALYSIS_SHEET_NAME As String = "Analysis"
Private Const TRANSLATIONS_SHEET_NAME As String = "Translations"

Private Type TSetupImportState
    importPath As String                    'Path to the source setup workbook
    progressDisplay As Object               'UI element displaying progress messages
    importWorkbook As Workbook              'Lazy loaded reference to the imported workbook
    appScope As IApplicationState           'Application state guard ensuring Excel settings are restored
End Type

Private this As TSetupImportState


'@section Factory helpers
'===============================================================================
'@sub-title Create a ready-to-use service instance
'@param importPath String path to the source setup workbook.
'@param progressDisplay Object receiving textual updates (must expose Caption).
'@return ISetupImportService configured with the provided path and progress display.
Public Function Create(ByVal importPath As String, ByVal progressDisplay As Object) As ISetupImportService
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
        scope.ApplyBusyState
    End If
End Sub

'@sub-title Restore Excel configuration when work completes.
Private Sub LeaveBusyState(Optional ByVal silent As Boolean = False)
    If this.appScope Is Nothing Then Exit Sub

    On Error GoTo RestoreFailed
        this.appScope.Restore
    On Error GoTo 0
    Exit Sub

RestoreFailed:
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
    Dim excludeNames As BetterArray
    Dim sheetUnlocked As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set workbook = ThisWorkbook
    ValidatePasswords pass
    Set sheetsList = EnsureSheetsList(sheetsList)
    Set excludeNames = New BetterArray

    EnterBusyState
    On Error GoTo CleanFailed

    For idx = sheetsList.LowerBound To sheetsList.UpperBound
        sheetName = CStr(sheetsList.Item(idx))
        Set targetSheet = workbook.Worksheets(sheetName)

        pass.UnProtect sheetName
        sheetUnlocked = True

        CleanWorksheetTables targetSheet, excludeNames
        ClearWorksheetComments targetSheet

        ProtectWorksheet pass, sheetName
        sheetUnlocked = False
    Next idx

CleanCleanup:
    LeaveBusyState
    Set excludeNames = Nothing
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

    ValidatePasswords pass
    Set sheetsList = EnsureSheetsList(sheetsList)

    EnterBusyState
    On Error GoTo ImportFailed

        PrepareImport pass, sheetsList
        Set importBook = EnsureImportWorkbook()
        Set hostWorkbook = ThisWorkbook

        For idx = sheetsList.LowerBound To sheetsList.UpperBound
            sheetName = CStr(sheetsList.Item(idx))

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
                        hostTable.Import importTable, keepImportHeaders:=(sheetName = TRANSLATIONS_SHEET_NAME)
                    End If
                Next importList

                ProtectWorksheet pass, sheetName
                sheetUnlocked = False
            End If
        Next idx

ImportCleanup:
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
        CloseImportWorkbook
    On Error GoTo 0
    LeaveBusyState True
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
End Sub


'@section Preparation helpers
'===============================================================================
'@sub-title Ensure the exports and dictionary sheets are aligned prior to import.
Private Sub PrepareImport(ByVal pass As IPasswords, ByVal sheetsList As BetterArray)
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
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If sheetsList Is Nothing Then Exit Sub
    If Not sheetsList.Includes(EXPORT_SHEET_NAME) Then Exit Sub

    Set importBook = EnsureImportWorkbook()
    Set hostWorkbook = ThisWorkbook

    On Error GoTo PrepareFailed

    If WorksheetExists(importBook, EXPORT_SHEET_NAME) And WorksheetExists(importBook, DICTIONARY_SHEET_NAME) Then
        Set currentExportsSheet = hostWorkbook.Worksheets(EXPORT_SHEET_NAME)
        Set importExportsSheet = importBook.Worksheets(EXPORT_SHEET_NAME)
        Set currentExportObj = LLExport.Create(currentExportsSheet, 4, 1)
        Set importExportObj = LLExport.Create(importExportsSheet, 4, 1)

        If currentExportObj.NumberOfExports < importExportObj.NumberOfExports Then
            pass.UnProtect DICTIONARY_SHEET_NAME
            dictUnlocked = True
            pass.UnProtect EXPORT_SHEET_NAME
            exportsUnlocked = True

            Set currentDictionary = LLdictionary.Create(hostWorkbook.Worksheets(DICTIONARY_SHEET_NAME), 5, 1)

            Do While currentExportObj.NumberOfExports < importExportObj.NumberOfExports
                previousExports = currentExportObj.NumberOfExports
                currentExportObj.AddRows
                updatedExports = currentExportObj.NumberOfExports

                For counter = previousExports + 1 To updatedExports
                    currentDictionary.InsertColumn "Export " & CStr(counter), After:="Export " & CStr(counter - 1)
                Next counter
            Loop
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
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

PrepareFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume PrepareCleanup
End Sub


'@section Validation helpers
'===============================================================================
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
    If progressDisplay Is Nothing Then
        ThrowError ProjectError.InvalidArgument, "Progress display object cannot be Nothing."
    End If

    On Error GoTo InvalidDisplay
        progressDisplay.Caption = progressDisplay.Caption
    On Error GoTo 0
    Exit Sub

InvalidDisplay:
    Err.Clear
    ThrowError ProjectError.InvalidArgument, "Progress display object must expose a Caption property."
End Sub

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

'@sub-title Close and release the imported workbook.
Private Sub CloseImportWorkbook()
    CloseWorkbook this.importWorkbook
    Set this.importWorkbook = Nothing
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
Private Sub CleanWorksheetTables(ByVal targetSheet As Worksheet, ByVal excludeNames As BetterArray)
    Dim listRef As ListObject
    Dim table As ICustomTable

    For Each listRef In targetSheet.ListObjects
        Set table = CustomTable.Create(listRef)
        table.Clean excludeNames:=excludeNames
        excludeNames.Clear
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
        usedCells.ClearCommentsThreaded
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
    If this.progressDisplay Is Nothing Then Exit Sub

    On Error Resume Next
        this.progressDisplay.Caption = message
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
