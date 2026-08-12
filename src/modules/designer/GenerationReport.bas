Attribute VB_Name = "GenerationReport"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Incremental generation report: harvests Checking objects and flushes them to the designer __check worksheet.")
'@depends CheckingOutput, Checking, HiddenNames, BetterArray, LinelistSpecs, LLdictionary, LLChoices, LLExport, Passwords, LLFormat
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@description
'Provides helpers for collecting diagnostic checkings from specification
'objects and flushing them incrementally to a __check worksheet on the
'designer workbook. Each flush appends to the existing output so the user
'can see progress as each generation phase completes.

Private Const SHEET_CHECKING As String = "__check"

'Render marker CheckingOutput keeps on its worksheet (ROW_MARKER_NAME
'there). InitReport drops it so a cleared sheet is formatted as a first
'render on every run.
Private Const ROW_MARKER_NAME As String = "CheckingOutputStartingRow"

'Module-level reference kept alive across flushes within one generation run.
Private reportOutput As CheckingOutput


'@section Public API
'===============================================================================

'@sub-title Initialise the report for a new generation run
'@details
'Resolves (or creates) the __check worksheet on the designer workbook,
'clears it, drops the render marker of the previous run, and prepares a
'fresh CheckingOutput writer. Must be called once at the start of
'clickGenerate before any FlushCheckings calls.
'@param designerBook Workbook. The designer workbook (ThisWorkbook).
Public Sub InitReport(ByVal designerBook As Workbook)
    Dim sh As Worksheet
    Dim store As HiddenNames

    Set sh = ResolveCheckingSheet(designerBook)
    sh.Cells.Clear

    'The marker survives a Cells.Clear. Keeping it makes the writer resume
    'on the cleared sheet and skip the first-render formatting.
    Set store = HiddenNames.Create(sh)
    If store.HasName(ROW_MARKER_NAME) Then store.RemoveName ROW_MARKER_NAME

    Set reportOutput = CheckingOutput.Create(sh, "Generation Report")
End Sub

'@sub-title Flush a batch of Checking objects to the report worksheet
'@details
'Appends the supplied checkings to the designer __check sheet.
'Skips silently when the batch is empty or InitReport was not called.
'@param checkBatch BetterArray. Collection of Checking instances to write.
Public Sub FlushCheckings(ByVal checkBatch As BetterArray)
    If reportOutput Is Nothing Then Exit Sub
    If checkBatch Is Nothing Then Exit Sub
    If checkBatch.Length = 0 Then Exit Sub

    reportOutput.PrintOutput checkBatch
End Sub

'@sub-title Harvest checkings from all specification collaborators after Prepare
'@details
'Collects Checking objects from Dictionary, Choices, Exports, Analysis,
'Passwords, and DesignFormat into a BetterArray ready for FlushCheckings.
'@param specs LinelistSpecs. The specs object after Prepare.
'@return BetterArray. Collection of Checking instances (may be empty).
Public Function HarvestSpecsCheckings(ByVal specs As LinelistSpecs) As BetterArray
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

    'The specifications themselves. A geobase the user asked for that would not
    'open is filed here, and it used to leave the linelist with no geographic
    'data and nothing said.
    If specs.HasCheckings Then
        result.Push specs.CheckingValues
    End If

    Set HarvestSpecsCheckings = result
End Function

'@sub-title Install the worksheet change handler for interactive filtering
'@details
'Call once after all flushes are complete so the user can filter the
'report by severity or title. Makes the report sheet visible and brings
'it to the front: the sheet is created very hidden, and Activate raises
'on a hidden sheet, so the visibility write comes first.
Public Sub FinaliseReport()
    If reportOutput Is Nothing Then Exit Sub
    reportOutput.EnsureWorksheetChangeHandler

    'Show the report so the user can read it
    On Error Resume Next
    reportOutput.Wksh.Visible = xlSheetVisible
    reportOutput.Wksh.Activate
    On Error GoTo 0

    'Release the module-level reference
    Set reportOutput = Nothing
End Sub


'@section Internal Helpers
'===============================================================================

'@description Resolve the __check worksheet, creating it if absent.
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
