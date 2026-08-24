Attribute VB_Name = "ImportChecking"
Option Explicit

'@Folder("Linelist")
'@ModuleDescription("The worksheet an import writes what it found onto.")
'@depends CheckingOutput, Checking, BetterArray, Passwords
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@description
'An import used to say "finished" and nothing else. Every step that decided to
'do nothing did it in silence: a dropdown that matched nothing, a variable the
'file had no column for, a choice list the two files disagreed about.
'
'This module writes all of it onto one worksheet of the linelist, in the same
'shape the generation report uses in the designer.
'
'THE SHEET IS WRITTEN, NOT SHOWN
'-------------------------------------------------------------------------------
'An import used to end by switching the screen onto this sheet, which put a
'wall of text in front of a user who had asked for none of it. The write is
'silent now: the sheet is very hidden and stays that way, and the import says
'nothing about it. The user reaches it from the button on the import report
'form, through ShowReportSheet, and puts it away with the close button of the
'ribbon, which knows this sheet by the name ReportSheetName answers.
'
'THE WRITE COMES BEFORE THE FORM
'-------------------------------------------------------------------------------
'Two subs, and they do one thing each: WriteImportCheckings writes the sheet,
'ShowReportSheet shows it. An import calls the write, and it must run BEFORE the
'import report form is shown, because the Open Log button of that form calls
'ShowReportSheet. A sheet written after the form was dismissed is a button that
'does nothing.
'
'THE WORKBOOK IS UNPROTECTED AROUND EVERY VISIBILITY WRITE
'-------------------------------------------------------------------------------
'The workbook structure guards sheet visibility, so adding this sheet, hiding
'it and showing it all need the workbook open. The worksheet ITSELF is never
'protected: the report carries filter dropdowns the user types into.

Private Const SHEET_IMPORT_CHECKING As String = "__import_checking"


'@section Public API
'===============================================================================

'@description The name of the worksheet an import writes onto. The sheet
'carries no sheet tag, so its name is the only mark it has, and the ribbon
'close button reads it from here rather than spelling it a second time.
'@return String. The worksheet name.
Public Function ReportSheetName() As String
    ReportSheetName = SHEET_IMPORT_CHECKING
End Function


'@description Write what one import found onto the import checking worksheet.
'The sheet is emptied first, so it always describes the last import. Nothing is
'shown and nothing is asked: the sheet stays very hidden until the user opens
'it from the import report form.
'@param sourceWkb Workbook. The linelist workbook.
'@param checks Checking. What the import filed.
Public Sub WriteImportCheckings(ByVal sourceWkb As Workbook, ByVal checks As Checking)

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

    'CheckingOutput writes a hidden sheet as readily as a shown one, so the
    'report is complete on a sheet nobody has looked at yet.
    writer.PrintOutput batch

    'The filter dropdowns need a handler in the worksheet code module, which
    'needs trust access to the VBA project. A linelist without that trust still
    'gets the report, without the filtering.
    On Error Resume Next
    writer.EnsureWorksheetChangeHandler
    On Error GoTo 0
End Sub


'@description Put the import report worksheet on screen and land the user on
'it. The workbook is opened around the visibility write, because the workbook
'structure guards it, and closed again on both paths. The WORKSHEET is left
'unprotected: its filter dropdowns are cells the user types into.
'A workbook that never ran an import has no sheet to show and the walk leaves
'quietly.
'@param sourceWkb Workbook. The linelist workbook.
'@return Boolean. True when the sheet is on screen.
Public Function ShowReportSheet(ByVal sourceWkb As Workbook) As Boolean

    Dim sh As Worksheet
    Dim pass As Passwords

    Set sh = ExistingCheckingSheet(sourceWkb)
    If sh Is Nothing Then Exit Function

    Set pass = PasswordManagerOf()
    If Not pass Is Nothing Then pass.UnProtect sourceWkb

    On Error Resume Next
    sh.Visible = xlSheetVisible
    sh.Activate
    On Error GoTo 0

    If Not pass Is Nothing Then pass.Protect sourceWkb

    ShowReportSheet = (sh.Visible = xlSheetVisible)
End Function


'@section Internal Helpers
'===============================================================================

'@description The password manager the event service holds, so the add and the
'unhide below can open the workbook. A workbook with no usable keys answers
'Nothing and the caller runs bare rather than ending the walk: a report the
'user cannot be shown must not take down the import that produced it.
'@return Passwords. The manager, or Nothing.
Private Function PasswordManagerOf() As Passwords
    Dim linelistEvents As EventLinelist

    On Error Resume Next
    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If Not linelistEvents Is Nothing Then _
        Set PasswordManagerOf = linelistEvents.PasswordManager()
    On Error GoTo 0
End Function


'@description The import checking worksheet of a workbook, or Nothing when the
'workbook has none. A workbook that never ran an import is the ordinary case.
'@param wb Workbook. The linelist workbook.
'@return Worksheet. The sheet, or Nothing.
Private Function ExistingCheckingSheet(ByVal wb As Workbook) As Worksheet
    'The worksheets collection raises 9 on a missing name, and the miss is the
    'answer rather than a fault.
    On Error Resume Next
    Set ExistingCheckingSheet = wb.Worksheets(SHEET_IMPORT_CHECKING)
    On Error GoTo 0
End Function


'@description Resolve the import checking worksheet, creating it very hidden
'when absent. The sheet is appended, because a bare Add puts a new sheet in
'front of whatever is active and this one belongs at the end.
'The workbook is opened around the add and the hide, because the workbook
'structure guards both, and closed again either way. A refusal leaves the sheet
'wherever it is: the report is still written onto it, and a report the user
'cannot be shown must never take down the import that produced it.
'@param wb Workbook. The linelist workbook.
'@return Worksheet. The resolved or newly created worksheet.
Private Function ResolveCheckingSheet(ByVal wb As Workbook) As Worksheet

    Dim sh As Worksheet
    Dim pass As Passwords

    Set sh = ExistingCheckingSheet(wb)
    If Not sh Is Nothing Then
        Set ResolveCheckingSheet = sh
        Exit Function
    End If

    Set pass = PasswordManagerOf()
    If Not pass Is Nothing Then pass.UnProtect wb

    On Error Resume Next
    Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    If Not sh Is Nothing Then
        sh.Name = SHEET_IMPORT_CHECKING
        'Born hidden. The import says nothing about this sheet, so it must not
        'appear in the tab bar of a user who never asked for it.
        sh.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    If Not pass Is Nothing Then pass.Protect wb

    Set ResolveCheckingSheet = sh
End Function
