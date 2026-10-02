VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseWorksheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides safe worksheet removal utilities for disease setup workflows"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Provides safe worksheet removal utilities for disease setup workflows")
'@depends ApplicationState
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'Removes disease worksheets safely and tells whether one exists.

'@section Factory
'===============================================================================

'@sub-title Build a fresh worksheet manager.
'@return DiseaseWorksheetManager ready to use.
Public Function Create() As DiseaseWorksheetManager
    Set Create = New DiseaseWorksheetManager
End Function

'@section Commands
'===============================================================================

'@sub-title Remove the specified worksheet using guarded Application settings.
'@param workbook Workbook hosting the worksheet collection.
'@param sheetName String worksheet name to delete.
'@return Boolean True when the worksheet existed and has been removed.
Public Function RemoveWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Boolean
    Dim guard As ApplicationState
    Dim targetSheet As Worksheet
    Dim errMessage As String

    ValidateWorkbook targetBook
    ValidateSheetName sheetName

    Set targetSheet = FindWorksheet(targetBook, sheetName)
    If targetSheet Is Nothing Then Exit Function

    On Error GoTo Fail

    Set guard = CreateGuard(targetBook.Application)
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
Public Function WorksheetExists(ByVal targetBook As Workbook, ByVal sheetName As String) As Boolean
    ValidateWorkbook targetBook
    ValidateSheetName sheetName
    WorksheetExists = Not FindWorksheet(targetBook, sheetName) Is Nothing
End Function

'@section Internal Helpers
'===============================================================================

'@sub-title Ensure the workbook reference is valid.
Private Sub ValidateWorkbook(ByVal targetBook As Workbook)
    If targetBook Is Nothing Then
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
Private Function FindWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    Dim candidate As Worksheet

    For Each candidate In targetBook.Worksheets
        If StrComp(candidate.Name, sheetName, vbTextCompare) = 0 Then
            Set FindWorksheet = candidate
            Exit Function
        End If
    Next
End Function

'@sub-title Instantiate and activate an Application guard.
Private Function CreateGuard(ByVal excelApp As Application) As ApplicationState
    Dim guard As ApplicationState

    Set guard = ApplicationState.Create(excelApp)
    guard.ApplyBusyState suppressEvents:=True
    Set CreateGuard = guard
End Function

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseWorksheetManager", errMessage
End Sub
