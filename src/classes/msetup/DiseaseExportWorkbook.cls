VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseExportWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides a safe lifecycle for export workbooks used by disease workflows"

'@PredeclaredId
'@Folder("Msetup")
'@ModuleDescription("Provides a safe lifecycle for export workbooks used by disease workflows")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'DiseaseExportWorkbook.cls without interface -- single implementation, no test
'fake, no compile cycle.

'@section Types
'===============================================================================

Private Type TExportWorkbook
    Book As Excel.Workbook
End Type

Private this As TExportWorkbook

'@section Factory
'===============================================================================

'@sub-title Build a fresh manager holding no workbook yet.
'@return DiseaseExportWorkbook ready to hand out a workbook on demand.
Public Function Create() As DiseaseExportWorkbook
    Set Create = New DiseaseExportWorkbook
End Function

'@section Lifecycle
'===============================================================================

'@sub-title Ensure a workbook instance is available for exports.
'@param excelApp Application host used to create the workbook.
'@return Workbook ready for export operations.
Public Function EnsureWorkbook(ByVal excelApp As Application) As Excel.Workbook
    If excelApp Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Application reference cannot be Nothing"
    End If

    If Not HasWorkbook Then
        Set this.Book = excelApp.Workbooks.Add
    End If

    Set EnsureWorkbook = this.Book
End Function

'@sub-title Persist the workbook to disk.
'@param filePath String destination path.
'@param fileFormat Optional Long holding an XlFileFormat value. Defaults to xlOpenXMLWorkbook.
'@param createBackup Optional Boolean. Defaults to False.
Public Sub SaveAs(ByVal filePath As String, _
                  Optional ByVal fileFormat As Long = xlOpenXMLWorkbook, _
                  Optional ByVal createBackup As Boolean = False)

    If LenB(filePath) = 0 Then
        ThrowError ProjectError.InvalidArgument, "File path cannot be empty"
    End If

    If Not HasWorkbook Then
        ThrowError ProjectError.ObjectNotInitialized, "No workbook is available to save"
    End If

    this.Book.SaveAs fileName:=filePath, fileFormat:=fileFormat, createBackup:=createBackup
End Sub

'@sub-title Close the workbook and clear the reference.
'@details Safe to call with nothing managed: that is a harmless no-op.
'@param saveChanges Optional Boolean controlling the SaveChanges flag.
Public Sub CloseWorkbook(Optional ByVal saveChanges As Boolean = False)
    If Not HasWorkbook Then
        Set this.Book = Nothing
        Exit Sub
    End If

    this.Book.Close saveChanges:=saveChanges
    Set this.Book = Nothing
End Sub

'@sub-title Release the workbook without saving, ensuring the reference is cleared.
'@details Never raises: a dead or missing workbook still ends released.
Public Sub Release()
    If Not this.Book Is Nothing Then
        ' The workbook may already be closed by the user; closing it again
        ' raises, and Release promises to clear the reference regardless.
        On Error Resume Next
        this.Book.Close saveChanges:=False
        On Error GoTo 0
    End If

    Set this.Book = Nothing
End Sub

'@section State
'===============================================================================

'@sub-title Indicates whether a workbook is currently managed.
Public Property Get HasWorkbook() As Boolean
    If this.Book Is Nothing Then Exit Property

    ' A workbook closed behind our back leaves a dead reference whose Name
    ' raises; treat that as no workbook. Scoped probe, restored right after.
    Dim probeName As String
    On Error Resume Next
    probeName = this.Book.Name
    On Error GoTo 0

    HasWorkbook = (LenB(probeName) > 0)
End Property

'@sub-title Retrieve the managed workbook reference.
Public Property Get Workbook() As Excel.Workbook
    Set Workbook = this.Book
End Property

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised error raising aligned with ProjectError codes.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseExportWorkbook", errMessage
End Sub
