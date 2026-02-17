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

'@class CodeTransfer
'@description
'Transfers VBA code components (classes, modules, forms, workbook code) from a
'source workbook to a target workbook using temporary file export/import cycles.
'Each component is exported to a temp file via the ITemporaryRepos, imported
'into the target, and the temp file is cleaned up. Consumers create an instance
'via Create and then call the transfer methods for each component type.
'@depends ITemporaryRepos
'@version 1.0 (2026-02-09)



Private Const CLASS_NAME As String = "CodeTransfer"

Private Type TCodeTransfer
    sourceWkb As Workbook
    targetWkb As Workbook
    tempRepos As ITemporaryRepos
End Type

Private this As TCodeTransfer


'@section Instantiation
'===============================================================================
'@description Factory method and Friend bindings for construction.

'@label:create
'@sub-title Create a CodeTransfer instance
'@details
'Creates a new CodeTransfer wired to transfer components from sourceWkb to
'targetWkb using tempRepos for intermediate file storage. All three parameters
'are validated as non-Nothing.
'@param sourceWkb Workbook. The workbook containing the VBA components to export.
'@param targetWkb Workbook. The workbook that will receive the imported components.
'@param tempRepos ITemporaryRepos. The temporary file repository for export/import.
'@return ICodeTransfer. A fully initialised instance ready for transfers.
'@throws ProjectError.ObjectNotInitialized When any parameter is Nothing.
'@export
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
'@description Friend setters used only during factory construction.

'@label:bin-source-wkb
'@prop-title Assign the source workbook
'@param value Workbook. The source workbook to store.
Friend Property Set BinSourceWkb(ByVal value As Workbook)
    Set this.sourceWkb = value
End Property

'@label:bin-target-wkb
'@prop-title Assign the target workbook
'@param value Workbook. The target workbook to store.
Friend Property Set BinTargetWkb(ByVal value As Workbook)
    Set this.targetWkb = value
End Property

'@label:bin-temp-repos
'@prop-title Assign the temporary repository
'@param value ITemporaryRepos. The temporary repository to store.
Friend Property Set BinTempRepos(ByVal value As ITemporaryRepos)
    Set this.tempRepos = value
End Property


'@section Operations
'===============================================================================
'@description Public methods for transferring VBA components.

'@label:transfer-class
'@sub-title Transfer a class module from source to target workbook
'@details
'Exports the named class module (.cls) from the source workbook to a temporary
'file named "CopieCls.cls", then imports it into the target workbook. The
'temporary file is cleaned up after import.
'@param className String. The VBA component name of the class to transfer.
'@export
Public Sub TransferClass(ByVal className As String)
    ExportImportComponent className, "CopieCls.cls"
End Sub

'@label:transfer-module
'@sub-title Transfer a standard module from source to target workbook
'@details
'Exports the named standard module (.bas) from the source workbook to a
'temporary file named "CopieMod.bas", then imports it into the target workbook.
'@param moduleName String. The VBA component name of the module to transfer.
'@export
Public Sub TransferModule(ByVal moduleName As String)
    ExportImportComponent moduleName, "CopieMod.bas"
End Sub

'@label:transfer-form
'@sub-title Transfer a UserForm from source to target workbook
'@details
'Exports the named UserForm (.frm/.frx) from the source workbook to temporary
'files, imports it into the target workbook, then deletes both the .frm and
'.frx temporary files.
'@param formName String. The VBA component name of the form to transfer.
'@export
Public Sub TransferForm(ByVal formName As String)
    Dim formPath As String
    Dim frxPath As String
    Dim vbComp As Object

    'Verify form exists before attempting export
    On Error Resume Next
    Set vbComp = this.sourceWkb.VBProject.VBComponents(formName)
    On Error GoTo 0

    If vbComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Form '" & formName & "' not found in source workbook"
    End If

    formPath = this.tempRepos.CreateFilePath("CopieUsf.frm")
    frxPath = this.tempRepos.CreateFilePath("CopieUsf.frx")

    this.tempRepos.DeleteFile formPath
    this.tempRepos.DeleteFile frxPath

    vbComp.Export formPath
    this.targetWkb.VBProject.VBComponents.Import formPath

    this.tempRepos.DeleteFile formPath
    this.tempRepos.DeleteFile frxPath
End Sub

'@label:transfer-workbook-code
'@sub-title Copy code from a source module into the target's ThisWorkbook component
'@details
'Reads all code lines from the named module in the source workbook and writes
'them into the target workbook's ThisWorkbook code module, replacing any
'existing content. Does not create a new component; instead overwrites the
'built-in ThisWorkbook code module.
'@param moduleName String. The source module whose code will be copied.
'@export
Public Sub TransferWorkbookCode(ByVal moduleName As String)
    Dim codeContent As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim sourceComp As Object

    'Verify source module exists
    On Error Resume Next
    Set sourceComp = this.sourceWkb.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    If sourceComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Module '" & moduleName & "' not found in source workbook"
    End If

    With sourceComp.CodeModule
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

'@label:copy-module-text
'@sub-title Copy code text from a source module to an existing target module
'@details
'Reads all code from the named module in the source workbook and writes it
'into the same-named module in the target workbook, replacing existing content.
'Used for template workbooks where the target module already exists.
'@param moduleName String. The module name (must exist in both workbooks).
'@export
Public Sub CopyModuleText(ByVal moduleName As String)
    Dim codeContent As String
    Dim codeMod As Object
    Dim sourceComp As Object
    Dim targetComp As Object

    'Verify module exists in both workbooks
    On Error Resume Next
    Set sourceComp = this.sourceWkb.VBProject.VBComponents(moduleName)
    Set targetComp = this.targetWkb.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    If sourceComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Module '" & moduleName & "' not found in source workbook"
    End If

    If targetComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Module '" & moduleName & "' not found in target workbook"
    End If

    With sourceComp.CodeModule
        codeContent = .Lines(1, .CountOfLines)
    End With

    Set codeMod = targetComp.CodeModule

    With codeMod
        .DeleteLines 1, .CountOfLines
        .AddFromString codeContent
    End With
End Sub

'@label:copy-module-code
'@sub-title Copy code between components within the target workbook
'@details
'Within the target workbook, reads code from the source component, writes it
'into the target component (replacing any existing content), then removes the
'source component from the VBProject. Useful after importing a module that
'should be injected into a different component (e.g., a worksheet module).
'@param sourceName String. The component to read code from and then remove.
'@param targetName String. The component to write code into.
'@throws ProjectError.ObjectNotInitialized When source or target component is
'not found in the target workbook.
'@export
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


'@section Helpers
'===============================================================================
'@description Private utilities for export/import and error handling.

'@label:export-import-component
'@sub-title Export a VBA component from source and import into target
'@details
'Exports the named component to a temporary file, imports it into the target
'workbook, then cleans up the temporary file. Used by TransferClass and
'TransferModule.
'@param componentName String. The VBA component name to transfer.
'@param tempFileName String. The temporary file name (e.g., "CopieCls.cls").
Private Sub ExportImportComponent(ByVal componentName As String, ByVal tempFileName As String)
    Dim tempPath As String
    Dim vbComp As Object

    'Verify component exists before attempting export
    On Error Resume Next
    Set vbComp = this.sourceWkb.VBProject.VBComponents(componentName)
    On Error GoTo 0

    If vbComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Component '" & componentName & "' not found in source workbook"
    End If

    tempPath = this.tempRepos.CreateFilePath(tempFileName)
    this.tempRepos.DeleteFile tempPath

    vbComp.Export tempPath
    this.targetWkb.VBProject.VBComponents.Import tempPath

    this.tempRepos.DeleteFile tempPath
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. The error code to raise.
'@param errorMessage String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise CLng(errNumber), CLASS_NAME, errorMessage
End Sub


'@section InterfaceImplementation
'===============================================================================
'@description Delegated members satisfying the ICodeTransfer contract.
'See the corresponding Public members above for full documentation.

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
