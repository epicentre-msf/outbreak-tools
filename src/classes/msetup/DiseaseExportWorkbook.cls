VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseExportWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Manages creation, saving, and disposal of disease export workbooks"

'@Folder("Msetup")
'@ModuleDescription("Manages creation, saving, and disposal of disease export workbooks")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseExportWorkbook

'@section State
'===============================================================================

Private Type TExportWorkbook
    ExcelApp As Application
    Workbook As Workbook
End Type

Private this As TExportWorkbook

'@section Lifecycle
'===============================================================================

'@sub-title Ensure a workbook is available and return it.
Public Function EnsureWorkbook(ByVal excelApp As Application) As Workbook
    ValidateApplication excelApp

    If Not this.Workbook Is Nothing Then
        Set EnsureWorkbook = this.Workbook
        Exit Function
    End If

    On Error GoTo Fail
        Set this.ExcelApp = excelApp
        Set this.Workbook = excelApp.Workbooks.Add
        Set EnsureWorkbook = this.Workbook
        Exit Function
Fail:
    ReleaseInternal
    ThrowError ProjectError.SomethingWentWrong, "Unable to create export workbook: " & Err.Description
End Function

'@sub-title Persist the workbook to the supplied path.
Public Sub SaveAs(ByVal filePath As String, _
                  Optional ByVal fileFormat As XlFileFormat = xlOpenXMLWorkbook, _
                  Optional ByVal createBackup As Boolean = False)

    ValidateWorkbookExists

    On Error GoTo Fail
        this.Workbook.SaveAs fileName:=filePath, _
                              FileFormat:=fileFormat, _
                              CreateBackup:=createBackup
        Exit Sub
Fail:
    ThrowError ProjectError.SomethingWentWrong, "Failed to save export workbook to '" & filePath & "': " & Err.Description
End Sub

'@sub-title Close the workbook and release references.
Public Sub Close(Optional ByVal saveChanges As Boolean = False)
    If this.Workbook Is Nothing Then Exit Sub

    On Error GoTo Fail
        this.Workbook.Close saveChanges:=saveChanges
        ReleaseInternal closeWorkbook:=False
        Exit Sub
Fail:
    ReleaseInternal closeWorkbook:=False
    ThrowError ProjectError.SomethingWentWrong, "Failed to close export workbook: " & Err.Description
End Sub

'@sub-title Release the workbook without saving changes.
Public Sub Release()
    ReleaseInternal closeWorkbook:=True
End Sub

'@section State
'===============================================================================

'@sub-title Expose whether a workbook is currently managed.
Public Property Get HasWorkbook() As Boolean
    HasWorkbook = Not this.Workbook Is Nothing
End Property

'@sub-title Provide the managed workbook reference.
Public Property Get Workbook() As Workbook
    Set Workbook = this.Workbook
End Property

'@section Internal Helpers
'===============================================================================

'@sub-title Validate the Application reference.
Private Sub ValidateApplication(ByVal excelApp As Application)
    If excelApp Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Application reference cannot be Nothing"
    End If
End Sub

'@sub-title Ensure a workbook has been created before manipulating it.
Private Sub ValidateWorkbookExists()
    If this.Workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Export workbook has not been created"
    End If
End Sub

'@sub-title Clear stored references, optionally closing the workbook.
Private Sub ReleaseInternal(Optional ByVal closeWorkbook As Boolean = True)
    On Error Resume Next
        If closeWorkbook And Not this.Workbook Is Nothing Then
            this.Workbook.Close saveChanges:=False
        End If
    On Error GoTo 0

    Set this.Workbook = Nothing
    Set this.ExcelApp = Nothing
End Sub

'@section ErrorHandling
'===============================================================================

'@sub-title Centralised ProjectError-based error raising.
Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseExportWorkbook", errMessage
End Sub

'@section ClassLifecycle
'===============================================================================

Private Sub Class_Terminate()
    ReleaseInternal closeWorkbook:=True
End Sub

'@section Interface
'===============================================================================

Private Function IDiseaseExportWorkbook_EnsureWorkbook(ByVal excelApp As Application) As Workbook
    Set IDiseaseExportWorkbook_EnsureWorkbook = EnsureWorkbook(excelApp)
End Function

Private Sub IDiseaseExportWorkbook_SaveAs(ByVal filePath As String, Optional ByVal fileFormat As XlFileFormat = xlOpenXMLWorkbook, Optional ByVal createBackup As Boolean = False)
    SaveAs filePath, fileFormat, createBackup
End Sub

Private Sub IDiseaseExportWorkbook_Close(Optional ByVal saveChanges As Boolean = False)
    Close saveChanges
End Sub

Private Sub IDiseaseExportWorkbook_Release()
    Release
End Sub

Private Property Get IDiseaseExportWorkbook_HasWorkbook() As Boolean
    IDiseaseExportWorkbook_HasWorkbook = HasWorkbook
End Property

Private Property Get IDiseaseExportWorkbook_Workbook() As Workbook
    Set IDiseaseExportWorkbook_Workbook = Workbook
End Property
