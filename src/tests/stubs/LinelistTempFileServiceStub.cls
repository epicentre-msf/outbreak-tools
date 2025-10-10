VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemporaryReposStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stubbed ITemporaryRepos for lifecycle manager tests")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ITemporaryRepos

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

Private Sub ITemporaryRepos_EnsureReady()
    this.wasReadied = True
End Sub

Private Sub ITemporaryRepos_Reset()
    this.resetCount = this.resetCount + 1
End Sub

Private Property Get ITemporaryRepos_RootPath() As String
    ITemporaryRepos_RootPath = "stub_path"
End Property

Private Function ITemporaryRepos_CreateFilePath(ByVal fileName As String) As String
    this.lastCreatedPath = fileName
    ITemporaryRepos_CreateFilePath = "stub_path/" & fileName
End Function

Private Sub ITemporaryRepos_DeleteFile(ByVal path As String)
    'No-op for stub
End Sub

Private Sub ITemporaryRepos_DeleteAll()
    this.deleteAllCount = this.deleteAllCount + 1
End Sub

