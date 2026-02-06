VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Transfers VBA code components between workbooks using a temporary repository"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Transfers VBA code components between workbooks using a temporary repository")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

Implements ICodeTransfer

Private Const CLASS_NAME As String = "CodeTransfer"

Private Type TCodeTransfer
    sourceWkb As Workbook
    targetWkb As Workbook
    tempRepos As ITemporaryRepos
End Type

Private this As TCodeTransfer

'@section Factory
'===============================================================================

Public Function Create(ByVal sourceWkb As Workbook, _
                       ByVal targetWkb As Workbook, _
                       ByVal tempRepos As ITemporaryRepos) As ICodeTransfer

    If sourceWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Source workbook is required"
    End If

    If targetWkb Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Target workbook is required"
    End If

    If tempRepos Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Temporary repository is required"
    End If

    Dim instance As CodeTransfer
    Set instance = New CodeTransfer

    Set instance.BinSourceWkb = sourceWkb
    Set instance.BinTargetWkb = targetWkb
    Set instance.BinTempRepos = tempRepos

    Set Create = instance
End Function

'@section Bindings
'===============================================================================

Friend Property Set BinSourceWkb(ByVal value As Workbook)
    Set this.sourceWkb = value
End Property

Friend Property Set BinTargetWkb(ByVal value As Workbook)
    Set this.targetWkb = value
End Property

Friend Property Set BinTempRepos(ByVal value As ITemporaryRepos)
    Set this.tempRepos = value
End Property

'@section Public methods
'===============================================================================

'Transfer a class module from source to target workbook via temp .cls file.
Public Sub TransferClass(ByVal className As String)
    ExportImportComponent className, "CopieCls.cls"
End Sub

'Transfer a standard module from source to target workbook via temp .bas file.
Public Sub TransferModule(ByVal moduleName As String)
    ExportImportComponent moduleName, "CopieMod.bas"
End Sub

'Transfer a UserForm from source to target workbook via temp .frm/.frx files.
Public Sub TransferForm(ByVal formName As String)
    Dim formPath As String
    Dim frxPath As String

    formPath = this.tempRepos.CreateFilePath("CopieUsf.frm")
    frxPath = this.tempRepos.CreateFilePath("CopieUsf.frx")

    this.tempRepos.DeleteFile formPath
    this.tempRepos.DeleteFile frxPath

    this.sourceWkb.VBProject.VBComponents(formName).Export formPath
    this.targetWkb.VBProject.VBComponents.Import formPath

    this.tempRepos.DeleteFile formPath
    this.tempRepos.DeleteFile frxPath
End Sub

'Copy code content from a source module into the target workbook's ThisWorkbook component.
Public Sub TransferWorkbookCode(ByVal moduleName As String)
    Dim codeContent As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object

    With this.sourceWkb.VBProject.VBComponents(moduleName).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    With this.targetWkb
        Set vbProj = .VBProject
        Set vbComp = vbProj.VBComponents(.codeName)
        Set codeMod = vbComp.CodeModule
    End With

    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'Copy code text from a source module to an existing same-named module in the target workbook.
'Used for template workbooks where the target module already exists.
Public Sub CopyModuleText(ByVal moduleName As String)
    Dim codeContent As String
    Dim codeMod As Object

    With this.sourceWkb.VBProject.VBComponents(moduleName).CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    Set codeMod = this.targetWkb.VBProject.VBComponents(moduleName).CodeModule

    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'Within the target workbook, copy code from one component to another, then remove the source.
'Useful after importing a module that should be injected into a different component (e.g. worksheet or ThisWorkbook).
Public Sub CopyModuleCode(ByVal sourceName As String, ByVal targetName As String)
    Dim vbProj As Object
    Dim sourceComponent As Object
    Dim targetComponent As Object
    Dim codeContent As String

    Set vbProj = this.targetWkb.VBProject

    On Error Resume Next
        Set sourceComponent = vbProj.VBComponents(sourceName)
        Set targetComponent = vbProj.VBComponents(targetName)
    On Error GoTo 0

    If sourceComponent Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Source component '" & sourceName & "' not found in target workbook"
    End If

    If targetComponent Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, _
                   "Target component '" & targetName & "' not found in target workbook"
    End If

    With sourceComponent.CodeModule
        If .CountOfLines > 0 Then
            codeContent = .Lines(1, .CountOfLines)
        Else
            codeContent = vbNullString
        End If
    End With

    With targetComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If LenB(codeContent) > 0 Then
            .AddFromString codeContent
        End If
    End With

    On Error Resume Next
        vbProj.VBComponents.Remove sourceComponent
    On Error GoTo 0
End Sub

'@section Private helpers
'===============================================================================

'Export a VBA component from source workbook, import into target workbook, clean up temp file.
Private Sub ExportImportComponent(ByVal componentName As String, ByVal tempFileName As String)
    Dim tempPath As String

    tempPath = this.tempRepos.CreateFilePath(tempFileName)
    this.tempRepos.DeleteFile tempPath

    this.sourceWkb.VBProject.VBComponents(componentName).Export tempPath
    this.targetWkb.VBProject.VBComponents.Import tempPath

    this.tempRepos.DeleteFile tempPath
End Sub

'Raise a project error with this class as the source.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise CLng(errNumber), CLASS_NAME, errorMessage
End Sub

'@section ICodeTransfer implementation
'===============================================================================

Private Sub ICodeTransfer_TransferClass(ByVal className As String)
    TransferClass className
End Sub

Private Sub ICodeTransfer_TransferModule(ByVal moduleName As String)
    TransferModule moduleName
End Sub

Private Sub ICodeTransfer_TransferForm(ByVal formName As String)
    TransferForm formName
End Sub

Private Sub ICodeTransfer_TransferWorkbookCode(ByVal moduleName As String)
    TransferWorkbookCode moduleName
End Sub

Private Sub ICodeTransfer_CopyModuleText(ByVal moduleName As String)
    CopyModuleText moduleName
End Sub

Private Sub ICodeTransfer_CopyModuleCode(ByVal sourceName As String, ByVal targetName As String)
    CopyModuleCode sourceName, targetName
End Sub
