VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseWorksheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides safe worksheet removal utilities for disease setup workflows"

'@Folder("Msetup")
'@ModuleDescription("Provides safe worksheet removal utilities for disease setup workflows")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseWorksheetManager

'@section Commands
'===============================================================================

'@sub-title Remove the specified worksheet using guarded Application settings.
'@param workbook Workbook hosting the worksheet collection.
'@param sheetName String worksheet name to delete.
'@return Boolean True when the worksheet existed and has been removed.
Public Function RemoveWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
    Dim guard As IDiseaseApplicationState
    Dim targetSheet As Worksheet
    Dim errMessage As String

    ValidateWorkbook workbook
    ValidateSheetName sheetName

    Set targetSheet = FindWorksheet(workbook, sheetName)
    If targetSheet Is Nothing Then Exit Function

    On Error GoTo Fail

    Set guard = CreateGuard(workbook.Application)
    targetSheet.Delete
    RemoveWorksheet = True

    On Error GoTo 0

CleanExit:
    If Not guard Is Nothing Then guard.Restore
    Exit Function

Fail:
    errMessage = Err.Description
    On Error Resume Next
        If Not guard Is Nothing Then guard.Restore
    On Error GoTo 0
    ThrowError ProjectError.SomethingWentWrong, "Failed to remove worksheet '" & sheetName & "': " & errMessage
End Function

'@section Queries
'===============================================================================

'@sub-title Test whether a worksheet exists within the workbook.
'@param workbook Workbook hosting the worksheet collection.
'@param sheetName String worksheet name to locate.
'@return Boolean True when the worksheet exists.
Public Function WorksheetExists(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
    ValidateWorkbook workbook
    ValidateSheetName sheetName
    WorksheetExists = Not FindWorksheet(workbook, sheetName) Is Nothing
End Function

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the workbook reference is valid.
Private Sub ValidateWorkbook(ByVal workbook As Workbook)
    If workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook reference cannot be Nothing"
    End If
End Sub

'@sub-title Ensure the worksheet name is not empty.
Private Sub ValidateSheetName(ByVal sheetName As String)
    If LenB(sheetName) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Worksheet name cannot be empty"
    End If
End Sub

'@sub-title Locate the worksheet by name using case-insensitive comparison.
Private Function FindWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In workbook.Worksheets
        If StrComp(candidate.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = candidate
            Exit Function
        End If
    Next
End Function

'@sub-title Instantiate and activate an Application guard.
Private Function CreateGuard(ByVal excelApp As Application) As IDiseaseApplicationState
    Dim guard As IDiseaseApplicationState

    Set guard = New DiseaseApplicationState
    guard.BeginGuard excelApp, False
    Set CreateGuard = guard
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseWorksheetManager", errMessage
End Sub

'@section Interface
'===============================================================================

Private Function IDiseaseWorksheetManager_RemoveWorksheet(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
    IDiseaseWorksheetManager_RemoveWorksheet = RemoveWorksheet(workbook, sheetName)
End Function

Private Function IDiseaseWorksheetManager_WorksheetExists(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
    IDiseaseWorksheetManager_WorksheetExists = WorksheetExists(workbook, sheetName)
End Function
