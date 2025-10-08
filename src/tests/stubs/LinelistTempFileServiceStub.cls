VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistTempFileServiceStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stubbed ILinelistTempFileService for lifecycle manager tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ILinelistTempFileService

Private Type TTempService
    wasReadied As Boolean
    resetCount As Long
    deleteAllCount As Long
    lastCreatedPath As String
End Type

Private this As TTempService

Public Property Get WasReadied() As Boolean
    WasReadied = this.wasReadied
End Property

Public Property Get ResetCount() As Long
    ResetCount = this.resetCount
End Property

Public Property Get DeleteAllCount() As Long
    DeleteAllCount = this.deleteAllCount
End Property

Public Property Get LastCreatedPath() As String
    LastCreatedPath = this.lastCreatedPath
End Property

Private Sub ILinelistTempFileService_EnsureReady()
    this.wasReadied = True
End Sub

Private Sub ILinelistTempFileService_Reset()
    this.resetCount = this.resetCount + 1
End Sub

Private Property Get ILinelistTempFileService_RootPath() As String
    ILinelistTempFileService_RootPath = "stub_path"
End Property

Private Function ILinelistTempFileService_CreateFilePath(ByVal fileName As String) As String
    this.lastCreatedPath = fileName
    ILinelistTempFileService_CreateFilePath = "stub_path/" & fileName
End Function

Private Sub ILinelistTempFileService_DeleteFile(ByVal path As String)
    'No-op for stub
End Sub

Private Sub ILinelistTempFileService_DeleteAll()
    this.deleteAllCount = this.deleteAllCount + 1
End Sub

