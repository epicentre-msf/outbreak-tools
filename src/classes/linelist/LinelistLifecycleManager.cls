VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistLifecycleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Handles workbook disposal and reset operations for the Linelist workflow"

'@Folder("Linelist")
'@ModuleDescription("Handles workbook disposal and reset operations for the Linelist workflow")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistLifecycleManager

'@section State and constants
'===============================================================================
Private Type TLifecycle
    accessor As ILinelistWorkbookAccessor
    releaser As ILinelistWorkbookReleaser
    tempFiles As ILinelistTempFileService
    stateScope As ILinelistApplicationStateScope
End Type

Private this As TLifecycle
Private Const CLASSNAME As String = "LinelistLifecycleManager"


'@section Factory helpers
'===============================================================================
Public Function Create(ByVal accessor As ILinelistWorkbookAccessor, _
                       ByVal tempFiles As ILinelistTempFileService, _
                       Optional ByVal stateScope As ILinelistApplicationStateScope, _
                       Optional ByVal releaser As ILinelistWorkbookReleaser) As ILinelistLifecycleManager

    Dim manager As LinelistLifecycleManager
    Set manager = New LinelistLifecycleManager
    manager.Initialise accessor, tempFiles, stateScope, releaser
    Set Create = manager
End Function

Friend Sub Initialise(ByVal accessor As ILinelistWorkbookAccessor, _
                      ByVal tempFiles As ILinelistTempFileService, _
                      ByVal stateScope As ILinelistApplicationStateScope, _
                      ByVal releaser As ILinelistWorkbookReleaser)

    EnsureAccessor accessor
    EnsureTempService tempFiles

    Set this.accessor = accessor
    Set this.tempFiles = tempFiles
    Set this.stateScope = stateScope

    If releaser Is Nothing Then
        Set this.releaser = TryCastReleaser(accessor)
    Else
        Set this.releaser = releaser
    End If
End Sub


'@section Lifecycle helpers
'===============================================================================
Private Sub ResetInternal(ByVal saveChanges As Boolean)
    Dim workbook As Workbook

    EnsureDependencies
    Set workbook = ResolveWorkbook

    CloseWorkbook workbook, saveChanges
    ClearWorkbookReference

    CleanupTemporaryFiles
    RefreshScope
End Sub

Private Sub CleanupTemporaryFiles()
    On Error GoTo CleanupFail
        this.tempFiles.DeleteAll
        this.tempFiles.Reset
        On Error GoTo 0
        Exit Sub

CleanupFail:
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Failed to clean temporary files: " & Err.Description
End Sub

Private Sub RefreshScope()
    If this.stateScope Is Nothing Then Exit Sub
    this.stateScope.RefreshSnapshot
End Sub

Private Sub CloseWorkbook(ByVal workbook As Workbook, ByVal saveChanges As Boolean)
    If workbook Is Nothing Then Exit Sub

    On Error GoTo CloseFail
        workbook.Close SaveChanges:=saveChanges
        On Error GoTo 0
        Exit Sub

CloseFail:
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Failed to close linelist workbook: " & Err.Description
End Sub

Private Sub ClearWorkbookReference()
    If Not this.releaser Is Nothing Then
        this.releaser.ClearOutputWorkbook
    End If
End Sub

Private Function ResolveWorkbook() As Workbook
    Dim workbook As Workbook

    Set workbook = this.accessor.OutputWorkbook
    Set ResolveWorkbook = workbook
End Function


'@section Validation helpers
'===============================================================================
Private Sub EnsureAccessor(ByVal accessor As ILinelistWorkbookAccessor)
    If accessor Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Workbook accessor cannot be Nothing"
    End If
End Sub

Private Sub EnsureTempService(ByVal tempFiles As ILinelistTempFileService)
    If tempFiles Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Temporary file service cannot be Nothing"
    End If
End Sub

Private Sub EnsureDependencies()
    If this.accessor Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Lifecycle manager was disposed"
    End If

    If this.tempFiles Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Temporary file service is not available"
    End If
End Sub

Private Function TryCastReleaser(ByVal accessor As ILinelistWorkbookAccessor) As ILinelistWorkbookReleaser
    Dim releaser As ILinelistWorkbookReleaser

    On Error GoTo Missing
        Set releaser = accessor
        On Error GoTo 0
        Set TryCastReleaser = releaser
        Exit Function

Missing:
    Err.Clear
    Set TryCastReleaser = Nothing
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistLifecycleManager_Reset(Optional ByVal saveChanges As Boolean = False)
    ResetInternal saveChanges
End Sub

Private Sub ILinelistLifecycleManager_Dispose(Optional ByVal saveChanges As Boolean = False)
    ResetInternal saveChanges

    Set this.accessor = Nothing
    Set this.releaser = Nothing
    Set this.tempFiles = Nothing
    Set this.stateScope = Nothing
End Sub

