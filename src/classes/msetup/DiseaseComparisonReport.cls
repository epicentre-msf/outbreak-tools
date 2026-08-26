VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseComparisonReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Writes a disease comparison onto the __compRep worksheet"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Writes a disease comparison onto the __compRep worksheet")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'DiseaseComparisonReport.cls without interface -- single implementation, no
'test fake, no compile cycle.

'@class DiseaseComparisonReport
'@description
'Takes the five Checking bundles a DiseaseComparison answers and renders them
'on the __compRep worksheet through CheckingOutput. The class holds the
'workbook and the sheet name; the sheet itself is resolved on every render.
'
'ONE REPORT AT A TIME
'-------------------------------------------------------------------------------
'The sheet is cleared before anything is written, so an old report never
'survives a new run. The next render restarts at the first output row and the
'sheet ends up carrying the last comparison alone.
'
'PROTECTION AND VISIBILITY
'-------------------------------------------------------------------------------
'__compRep sits outside the protection matrix of the master setup. The deploy
'step protects Variables, Choices and the translations sheet alone
'(EventsMasterSetupRibbon), so the clear needs no unprotect. That same step
'hides __compRep, so the sheet is made visible again before the render starts.

'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "DiseaseComparisonReport"
Private Const REPORT_SHEET_NAME As String = "__compRep"

'@section State
'===============================================================================

Private Type TReportState
    hostBook As Workbook
    sheetName As String
End Type

Private this As TReportState

'@section Factory
'===============================================================================

'@sub-title Build a writer bound to a workbook.
'@param hostBook Workbook. The master setup workbook carrying the report sheet.
'@param reportSheetName Optional String. Report worksheet. Defaults to "__compRep".
'@return DiseaseComparisonReport ready to render.
Public Function Create(ByVal hostBook As Workbook, _
                       Optional ByVal reportSheetName As String = REPORT_SHEET_NAME) As DiseaseComparisonReport
    Dim instance As DiseaseComparisonReport

    Set instance = New DiseaseComparisonReport
    instance.Initialise hostBook, reportSheetName
    Set Create = instance
End Function

Public Sub Initialise(ByVal hostBook As Workbook, ByVal reportSheetName As String)
    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "hostBook cannot be Nothing"
    End If

    Set this.hostBook = hostBook
    this.sheetName = Trim$(reportSheetName)
    If LenB(this.sheetName) = 0 Then this.sheetName = REPORT_SHEET_NAME
End Sub

'@section Accessors
'===============================================================================

'@prop-title Name of the worksheet the report is written on.
Public Property Get SheetName() As String
    SheetName = this.sheetName
End Property

'@prop-title The report worksheet.
'@details Raises ElementNotFound when the workbook carries no such sheet.
Public Property Get ReportSheet() As Worksheet
    Set ReportSheet = ResolveWorksheet(this.sheetName)
End Property

'@section Output
'===============================================================================

'@sub-title Render a comparison on the report worksheet.
'@details
'Clears the sheet, prints the five bundles the comparison answers, installs the
'worksheet change handler that drives the filter, and activates the sheet.
'@param comparison DiseaseComparison. The comparison to render.
'@throws ProjectError.ObjectNotInitialized When comparison is Nothing.
'@throws ProjectError.ElementNotFound When the workbook has no report sheet.
Public Sub PrintReport(ByVal comparison As DiseaseComparison)
    Dim checkOutput As CheckingOutput
    Dim targetSheet As Worksheet

    If comparison Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "A comparison is required to write the report"
    End If

    Set targetSheet = ResolveWorksheet(this.sheetName)

    'The sheet is shown before anything is written on it: the render autofits
    'the columns it writes, and a hidden worksheet refuses that.
    ShowReport targetSheet

    'An old report never survives a new run.
    targetSheet.Cells.Clear
    Set checkOutput = CheckingOutput.Create(targetSheet)

    'An earlier render may have left rows filtered out of sight.
    targetSheet.Cells.EntireRow.Hidden = False

    checkOutput.PrintOutput comparison.Reports

    'EnsureWorksheetChangeHandler writes into the VBA project and many field
    'machines block that. The report reads fine without the filter, so this one
    'step is allowed to fail quietly.
    On Error Resume Next
        checkOutput.EnsureWorksheetChangeHandler
    On Error GoTo 0

    targetSheet.Activate
End Sub

'@sub-title Build the comparison over two tables and render it in one call.
'@param firstTable ListObject. The disease table of the first disease.
'@param secondTable ListObject. The disease table of the second disease.
'@param firstName Optional String. Name shown for the first disease.
'@param secondName Optional String. Name shown for the second disease.
'@return DiseaseComparison. The comparison that reached the sheet.
Public Function PrintComparison(ByVal firstTable As ListObject, ByVal secondTable As ListObject, _
                                Optional ByVal firstName As String = "Disease 1", _
                                Optional ByVal secondName As String = "Disease 2") As DiseaseComparison
    Dim comparison As DiseaseComparison

    Set comparison = DiseaseComparison.Create(firstTable, secondTable, firstName, secondName)
    PrintReport comparison
    Set PrintComparison = comparison
End Function

'@section Worksheet Helpers
'===============================================================================

'@sub-title Make the report visible.
'@details
'The deploy step hides __compRep. A hidden worksheet refuses the column autofit
'the render does and refuses to be activated, and the user pressed the button to
'read the report, so the sheet is shown before the render starts.
Private Sub ShowReport(ByVal targetSheet As Worksheet)
    If targetSheet.Visible <> xlSheetVisible Then
        targetSheet.Visible = xlSheetVisible
    End If
End Sub

'@sub-title Find a worksheet of the host workbook by name.
Private Function ResolveWorksheet(ByVal lookupName As String) As Worksheet
    Dim target As Worksheet

    'A missing sheet raises on the lookup; the message below names it.
    On Error Resume Next
        Set target = this.hostBook.Worksheets(lookupName)
    On Error GoTo 0

    If target Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Worksheet '" & lookupName & "' not found in the master setup workbook."
    End If

    Set ResolveWorksheet = target
End Function

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, CLASS_NAME, errMessage
End Sub
