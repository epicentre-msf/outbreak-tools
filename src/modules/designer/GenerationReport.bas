Attribute VB_Name = "GenerationReport"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Incremental generation report: harvests IChecking objects and flushes them to the designer __checking worksheet.")
'@depends CheckingOutput, ICheckingOutput, Checking, IChecking, HiddenNames, IHiddenNames, BetterArray, ILinelistSpecs, ILLdictionary, ILLChoices, ILLExport, IAnalysis, IPasswords, ILLFormat
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@description
'Provides helpers for collecting diagnostic checkings from specification
'objects and flushing them incrementally to a __checking worksheet on the
'designer workbook. Each flush appends to the existing output so the user
'can see progress as each generation phase completes.

Private Const SHEET_CHECKING As String = "__checking"

'Module-level reference kept alive across flushes within one generation run.
Private reportOutput As ICheckingOutput


'@section Public API
'===============================================================================

'@sub-title Initialise the report for a new generation run
'@details
'Resolves (or creates) the __checking worksheet on the designer workbook,
'clears it, and prepares a fresh CheckingOutput writer. Must be called
'once at the start of clickGenerate before any FlushCheckings calls.
'@param designerBook Workbook. The designer workbook (ThisWorkbook).
Public Sub InitReport(ByVal designerBook As Workbook)
    Dim sh As Worksheet

    Set sh = ResolveCheckingSheet(designerBook)
    sh.Cells.Clear

    Set reportOutput = CheckingOutput.Create(sh, "Generation Report")
End Sub

'@sub-title Flush a batch of IChecking objects to the report worksheet
'@details
'Appends the supplied checkings to the designer __checking sheet.
'Skips silently when the batch is empty or InitReport was not called.
'@param checkBatch BetterArray. Collection of IChecking instances to write.
Public Sub FlushCheckings(ByVal checkBatch As BetterArray)
    If reportOutput Is Nothing Then Exit Sub
    If checkBatch Is Nothing Then Exit Sub
    If checkBatch.Length = 0 Then Exit Sub

    reportOutput.PrintOutput checkBatch
End Sub

'@sub-title Harvest checkings from all specification collaborators after Prepare
'@details
'Collects IChecking objects from Dictionary, Choices, Exports, Analysis,
'Passwords, and DesignFormat into a BetterArray ready for FlushCheckings.
'@param specs ILinelistSpecs. The specs object after Prepare.
'@return BetterArray. Collection of IChecking instances (may be empty).
Public Function HarvestSpecsCheckings(ByVal specs As ILinelistSpecs) As BetterArray
    Dim result As BetterArray
    Set result = New BetterArray
    result.LowerBound = 1

    'Dictionary
    If specs.Dictionary.HasCheckings Then
        result.Push specs.Dictionary.CheckingValues
    End If

    'Choices
    If specs.Choices.HasCheckings Then
        result.Push specs.Choices.CheckingValues
    End If

    'Exports
    If specs.ExportObject.HasCheckings Then
        result.Push specs.ExportObject.CheckingValues
    End If

    'Analysis
    If specs.AnalysisObject.HasCheckings Then
        result.Push specs.AnalysisObject.CheckingValues
    End If

    'Passwords
    If specs.Password.HasCheckings Then
        result.Push specs.Password.CheckingValues
    End If

    'Format
    If specs.DesignFormat.HasCheckings Then
        result.Push specs.DesignFormat.CheckingValues
    End If

    Set HarvestSpecsCheckings = result
End Function

'@sub-title Install the worksheet change handler for interactive filtering
'@details
'Call once after all flushes are complete so the user can filter the
'report by severity or title.
Public Sub FinaliseReport()
    If reportOutput Is Nothing Then Exit Sub
    reportOutput.EnsureWorksheetChangeHandler

    'Activate the checking sheet so the user sees the report
    On Error Resume Next
    reportOutput.Wksh.Activate
    On Error GoTo 0

    'Release the module-level reference
    Set reportOutput = Nothing
End Sub

'@sub-title Clean up module state without finalising
'@details
'Releases the module-level CheckingOutput reference. Use this in error
'paths where the report should be abandoned.
Public Sub ResetReport()
    Set reportOutput = Nothing
End Sub


'@section Internal Helpers
'===============================================================================

'@description Resolve the __checking worksheet, creating it if absent.
'@param wb Workbook. The workbook to search.
'@return Worksheet. The resolved or newly created worksheet.
Private Function ResolveCheckingSheet(ByVal wb As Workbook) As Worksheet
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = wb.Worksheets(SHEET_CHECKING)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = SHEET_CHECKING
        sh.Visible = xlSheetVeryHidden
    End If

    Set ResolveCheckingSheet = sh
End Function
