Attribute VB_Name = "ImportChecking"
Option Explicit

'@Folder("Linelist")
'@ModuleDescription("The worksheet an import writes what it found onto.")
'@depends CheckingOutput, Checking, BetterArray
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@description
'An import used to say "finished" and nothing else. Every step that decided to
'do nothing did it in silence: a dropdown that matched nothing, a variable the
'file had no column for, a choice list the two files disagreed about.
'
'This module writes all of it onto one worksheet of the linelist, in the same
'shape the generation report uses in the designer. The sheet is a plain
'worksheet and the user closes it when they are done reading.

Private Const SHEET_IMPORT_CHECKING As String = "__import_checking"


'@section Public API
'===============================================================================

'@description Write what one import found onto the import checking worksheet.
'The sheet is emptied first, so it always describes the last import.
'@param sourceWkb Workbook. The linelist workbook.
'@param checks Checking. What the import filed.
Public Sub ShowImportCheckings(ByVal sourceWkb As Workbook, ByVal checks As Checking)

    Dim sh As Worksheet
    Dim writer As CheckingOutput
    Dim batch As BetterArray

    If checks Is Nothing Then Exit Sub
    If checks.Length = 0 Then Exit Sub

    Set sh = ResolveCheckingSheet(sourceWkb)
    If sh Is Nothing Then Exit Sub

    sh.Cells.Clear

    Set writer = CheckingOutput.Create(sh, "Import report")

    Set batch = New BetterArray
    batch.LowerBound = 1
    batch.Push checks

    writer.PrintOutput batch

    'The filter dropdowns need a handler in the worksheet code module, which
    'needs trust access to the VBA project. A linelist without that trust still
    'gets the report, without the filtering.
    On Error Resume Next
    writer.EnsureWorksheetChangeHandler
    sh.Activate
    On Error GoTo 0
End Sub


'@section Internal Helpers
'===============================================================================

'@description Resolve the import checking worksheet, creating it when absent.
'The sheet is appended, because a bare Add puts a new sheet in front of whatever
'is active and this one belongs at the end.
'@param wb Workbook. The linelist workbook.
'@return Worksheet. The resolved or newly created worksheet.
Private Function ResolveCheckingSheet(ByVal wb As Workbook) As Worksheet

    Dim sh As Worksheet

    On Error Resume Next
    Set sh = wb.Worksheets(SHEET_IMPORT_CHECKING)
    On Error GoTo 0

    If sh Is Nothing Then
        On Error Resume Next
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If Not sh Is Nothing Then sh.Name = SHEET_IMPORT_CHECKING
        On Error GoTo 0
    End If

    If Not sh Is Nothing Then sh.Visible = xlSheetVisible

    Set ResolveCheckingSheet = sh
End Function
