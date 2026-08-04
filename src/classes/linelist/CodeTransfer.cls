VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Transfers VBA code components between workbooks using a temporary repository"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Transfers VBA code components between workbooks using a temporary repository")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

Option Explicit

'@class CodeTransfer
'@description
'Moves VBA code components -- classes, standard modules, forms and workbook
'code -- from a source workbook into a target workbook. A component is exported
'to a file in the temporary repository, imported into the target, and the file
'is deleted. TransferClass, TransferModule, TransferForm, TransferWorkbookCode
'and TransferWorksheetCode each move one component, and CopyModuleText copies
'the text of one module. Linelist.TransferAllCode is the one consumer, and it
'moves 39 components in a single generation.
'
'A COMPONENT THE TARGET ALREADY HOLDS
'-------------------------------------------------------------------------------
'VBComponents.Import stays quiet when the target already carries a component of
'that name: VBA imports the arriving one under a second name and keeps both
'copies, and the project stops compiling. Every transfer therefore removes the
'component that is already there before it imports, and files a warning through
'Checking naming it. The source copy is the one that survives, and a template
'author who wants their own code kept gives it another name. Owner decision,
'2026-07-31.
'
'THE VBA PROJECT HAS TO BE READABLE
'-------------------------------------------------------------------------------
'Every member of this class reads a workbook's VBProject. Excel refuses that
'read with error 1004 while "Trust access to the VBA project object model" is
'off, and that is how a fresh install ships. SourceProject and TargetProject
'raise ErrorUnexpectedState naming the setting. Component lookups happen after
'that read, where a Nothing means the component is missing.
'
'A CODE MODULE IS RESTORED WHEN A WRITE FAILS
'-------------------------------------------------------------------------------
'Writing code into a component empties it first. The four writers keep the text
'they are about to replace and put it back when the write raises, so a failure
'leaves the target holding what it held before.
'
'CHECKINGS NEED A HARVEST LINE
'-------------------------------------------------------------------------------
'Entries filed here leave through HasCheckings and CheckingValues. Linelist
'keeps the instance after Prepare, and clickGenerate flushes the entries into
'the generation report.
'@depends TemporaryRepos, Checking
'@version 1.1 (2026-07-31)



Private Const CLASS_NAME As String = "CodeTransfer"
Private Const TRUST_HINT As String = _
    "Excel refused to read the VBA project of the workbook '"
Private Const TRUST_ACTION As String = _
    "'. Turn on 'Trust access to the VBA project object model' in the Excel " & _
    "macro security settings and generate again."

Private Type TCodeTransfer
    sourceBook As Workbook
    targetBook As Workbook
    repos As TemporaryRepos
    internalChecks As Checking
    checkCounter As Long
    hasChecks As Boolean
    sealed As Boolean
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
'are validated as non-Nothing, and the instance is sealed before it is handed
'back, so the three bindings hold for its whole life.
'@param sourceWkb Workbook. The workbook containing the VBA components to export.
'@param targetWkb Workbook. The workbook that will receive the imported components.
'@param tempRepos TemporaryRepos. The temporary file repository for export/import.
'@return CodeTransfer. A fully initialised instance ready for transfers.
'@throws ProjectError.ObjectNotInitialized When any parameter is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook, _
                       ByVal targetWkb As Workbook, _
                       ByVal tempRepos As TemporaryRepos) As CodeTransfer

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

    Set instance.SourceWkb = sourceWkb
    Set instance.TargetWkb = targetWkb
    Set instance.TempRepos = tempRepos
    instance.Seal

    Set Create = instance
End Function

'@label:seal
'@sub-title Close the instance to creation-only writes
'@details
'Called by the factory once the three bindings are in place. Swapping the target
'workbook part way through a transfer would scatter components across two files,
'so the setters refuse a second write.
'@export
Public Sub Seal()
    this.sealed = True
End Sub


'@section Bindings
'===============================================================================
'@description Friend setters used only during factory construction.

'@label:source-wkb
'@prop-title Assign the source workbook
'@param value Workbook. The workbook the components are read from.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Set SourceWkb(ByVal value As Workbook)
    GuardNotSealed "SourceWkb"
    Set this.sourceBook = value
End Property

'@label:target-wkb
'@prop-title Assign the target workbook
'@param value Workbook. The workbook the components are written into.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Set TargetWkb(ByVal value As Workbook)
    GuardNotSealed "TargetWkb"
    Set this.targetBook = value
End Property

'@label:temp-repos
'@prop-title Assign the temporary repository
'@param value TemporaryRepos. The folder the export files are written to.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Friend Property Set TempRepos(ByVal value As TemporaryRepos)
    GuardNotSealed "TempRepos"
    Set this.repos = value
End Property


'@section Operations
'===============================================================================
'@description Public methods for transferring VBA components.

'@label:transfer-class
'@sub-title Transfer a class module from source to target workbook
'@details
'Exports the named class module to a file called after the component, imports it
'into the target workbook, and deletes the file. A class of that name already in
'the target is removed first and the removal is filed through Checking.
'@param className String. The VBA component name of the class to transfer.
'@throws ProjectError.ElementNotFound When the source workbook has no such class.
'@export
Public Sub TransferClass(ByVal className As String)
    ExportImportComponent className, className & ".cls"
End Sub

'@label:transfer-module
'@sub-title Transfer a standard module from source to target workbook
'@details
'Exports the named standard module to a file called after the component, imports
'it into the target workbook, and deletes the file. A module of that name already
'in the target is removed first and the removal is filed through Checking.
'@param moduleName String. The VBA component name of the module to transfer.
'@throws ProjectError.ElementNotFound When the source workbook has no such module.
'@export
Public Sub TransferModule(ByVal moduleName As String)
    ExportImportComponent moduleName, moduleName & ".bas"
End Sub

'@label:transfer-form
'@sub-title Transfer a UserForm from source to target workbook
'@details
'Exports the named UserForm to a .frm and a .frx called after the component,
'imports it into the target workbook, and deletes both files. A form of that name
'already in the target is removed first and the removal is filed through
'Checking. The .frx path is built by swapping the extension, which is what the
'VBIDE writes beside the .frm.
'@param formName String. The VBA component name of the form to transfer.
'@throws ProjectError.ElementNotFound When the source workbook has no such form.
'@export
Public Sub TransferForm(ByVal formName As String)
    Dim formPath As String
    Dim frxPath As String
    Dim vbComp As Object
    Dim targetProj As Object
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    Set vbComp = SourceComponent(formName)
    Set targetProj = TargetProject()

    formPath = this.repos.CreateFilePath(formName & ".frm")
    frxPath = this.repos.CreateFilePath(formName & ".frx")

    this.repos.DeleteFile formPath
    this.repos.DeleteFile frxPath

    On Error GoTo TransferFailed

    vbComp.Export formPath
    RemoveTargetComponent targetProj, formName
    targetProj.VBComponents.Import formPath

    this.repos.DeleteFile formPath
    this.repos.DeleteFile frxPath
    Exit Sub

TransferFailed:
    'The error is copied before any cleanup. DeleteFile runs its own
    'On Error Resume Next, which clears Err.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description

    this.repos.DeleteFile formPath
    this.repos.DeleteFile frxPath
    Err.Raise failedNumber, failedSource, failedDescription
End Sub

'@label:transfer-workbook-code
'@sub-title Copy code from a source module into the target's ThisWorkbook component
'@details
'Reads all code lines from the named module in the source workbook and writes
'them into the target workbook's ThisWorkbook code module, replacing what is
'there. The component is reached through Workbook.CodeName, which is where a
'workbook's own event code lives. The previous text is put back when the write
'raises.
'@param moduleName String. The source module whose code will be copied.
'@throws ProjectError.ElementNotFound When the source workbook has no such module.
'@export
Public Sub TransferWorkbookCode(ByVal moduleName As String)
    Dim codeContent As String
    Dim targetProj As Object
    Dim sourceComp As Object
    Dim codeMod As Object

    Set sourceComp = SourceComponent(moduleName)
    codeContent = ReadCode(sourceComp.CodeModule)

    Set targetProj = TargetProject()
    Set codeMod = targetProj.VBComponents(this.targetBook.CodeName).CodeModule

    ReplaceCode codeMod, codeContent
End Sub

'@label:transfer-worksheet-code
'@sub-title Copy code from a source module into a worksheet's code module
'@details
'Reads all code lines from the named module in the source workbook and writes
'them into the code module of the worksheet identified by targetSheetName in the
'target workbook, replacing what is there. The worksheet is looked up by its
'user-visible Name and its CodeName then locates the VBComponent, because the
'caller passes translated sheet names. Used to inject event handlers such as
'Worksheet_SelectionChange during linelist generation. The previous text is put
'back when the write raises.
'@param moduleName String. The source module whose code will be copied.
'@param targetSheetName String. The user-visible name of the worksheet in the
'target workbook whose code module will receive the code.
'@throws ProjectError.ElementNotFound When the source module does not exist.
'@throws ProjectError.ElementNotFound When the target worksheet does not exist.
'@export
Public Sub TransferWorksheetCode(ByVal moduleName As String, ByVal targetSheetName As String)
    Dim codeContent As String
    Dim sourceComp As Object
    Dim targetSh As Worksheet
    Dim targetProj As Object
    Dim codeMod As Object

    Set sourceComp = SourceComponent(moduleName)
    codeContent = ReadCode(sourceComp.CodeModule)

    'Find the target worksheet by its user-visible name
    On Error Resume Next
    Set targetSh = this.targetBook.Worksheets(targetSheetName)
    On Error GoTo 0

    If targetSh Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Worksheet '" & targetSheetName & "' not found in target workbook"
    End If

    Set targetProj = TargetProject()
    Set codeMod = targetProj.VBComponents(targetSh.CodeName).CodeModule

    ReplaceCode codeMod, codeContent
End Sub

'@label:copy-module-text
'@sub-title Copy code text from a source module to an existing target module
'@details
'Reads all code from the named module in the source workbook and writes it into
'the same-named module in the target workbook, replacing what is there. Both
'workbooks have to carry the module already: this member writes text into a
'component and creates none. Linelist calls it on the template branch, where the
'target workbook arrives with its own ribbon module. The previous text is put
'back when the write raises.
'@param moduleName String. The module name. It has to exist in both workbooks.
'@throws ProjectError.ElementNotFound When either workbook lacks the module.
'@export
Public Sub CopyModuleText(ByVal moduleName As String)
    Dim codeContent As String
    Dim sourceComp As Object
    Dim targetComp As Object

    Set sourceComp = SourceComponent(moduleName)
    Set targetComp = FindComponent(TargetProject(), moduleName)

    If targetComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Module '" & moduleName & "' not found in target workbook"
    End If

    codeContent = ReadCode(sourceComp.CodeModule)

    ReplaceCode targetComp.CodeModule, codeContent
End Sub


'@section Diagnostics
'===============================================================================
'@description The Checking entries this instance filed, and how a caller reads them.

'@label:has-checkings
'@prop-title Whether anything was filed during the transfers
'@return Boolean. True once one entry has been recorded.
'@export
Public Property Get HasCheckings() As Boolean
    HasCheckings = this.hasChecks
End Property

'@label:checking-values
'@prop-title The diagnostic log of the transfers
'@details
'Answers the internal Checking object holding every entry filed by this
'instance, and Nothing while none has been filed. GenerationReport reads it
'through Linelist.
'@return Object. A Checking instance, or Nothing.
'@export
Public Property Get CheckingValues() As Object
    If Not HasCheckings() Then Exit Property
    Set CheckingValues = this.internalChecks
End Property


'@section Helpers
'===============================================================================
'@description Private utilities for export/import, code writing and error handling.

'@label:export-import-component
'@sub-title Export a VBA component from source and import it into target
'@details
'Exports the named component to a temporary file, removes any component of that
'name from the target, imports the file and deletes it. The temporary file is
'deleted again when any of those steps raises. Used by TransferClass and
'TransferModule.
'@param componentName String. The VBA component name to transfer.
'@param tempFileName String. The temporary file name, extension included.
Private Sub ExportImportComponent(ByVal componentName As String, _
                                  ByVal tempFileName As String)
    Dim tempPath As String
    Dim vbComp As Object
    Dim targetProj As Object
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    Set vbComp = SourceComponent(componentName)
    Set targetProj = TargetProject()

    tempPath = this.repos.CreateFilePath(tempFileName)
    this.repos.DeleteFile tempPath

    On Error GoTo TransferFailed

    vbComp.Export tempPath
    RemoveTargetComponent targetProj, componentName
    targetProj.VBComponents.Import tempPath

    this.repos.DeleteFile tempPath
    Exit Sub

TransferFailed:
    'The error is copied before any cleanup. DeleteFile runs its own
    'On Error Resume Next, which clears Err.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description

    this.repos.DeleteFile tempPath
    Err.Raise failedNumber, failedSource, failedDescription
End Sub

'@label:remove-target-component
'@sub-title Take a same-named component out of the target before an import
'@details
'VBComponents.Import stays quiet on a name clash and leaves the target holding
'two copies under two names, which stops the project compiling. The copy that is
'there goes first, and the removal is filed as a warning so the generation report
'names it.
'@param targetProj Object. The VBProject of the target workbook.
'@param componentName String. The component about to arrive.
Private Sub RemoveTargetComponent(ByVal targetProj As Object, _
                                  ByVal componentName As String)
    Dim existing As Object

    Set existing = FindComponent(targetProj, componentName)
    If existing Is Nothing Then Exit Sub

    targetProj.VBComponents.Remove existing

    LogInfo "Component '" & componentName & "' was already in the workbook '" & _
            this.targetBook.Name & "'. The copy that was there is removed and " & _
            "the one from the designer takes its place.", checkingWarning
End Sub

'@label:source-component
'@sub-title Resolve a component of the source workbook
'@param componentName String. The component name to look for.
'@return Object. The VBComponent.
'@throws ProjectError.ElementNotFound When the source workbook has no such component.
Private Function SourceComponent(ByVal componentName As String) As Object
    Dim vbComp As Object

    Set vbComp = FindComponent(SourceProject(), componentName)

    If vbComp Is Nothing Then
        ThrowError ProjectError.ElementNotFound, _
                   "Component '" & componentName & "' not found in source workbook"
    End If

    Set SourceComponent = vbComp
End Function

'@label:find-component
'@sub-title Look a component up and answer Nothing when it is absent
'@param vbProj Object. The VBProject to search.
'@param componentName String. The component name to look for.
'@return Object. The VBComponent, or Nothing.
Private Function FindComponent(ByVal vbProj As Object, _
                               ByVal componentName As String) As Object
    'The suppression covers the lookup alone. The project itself is resolved
    'above, where a blocked project raises and names the setting to turn on.
    On Error Resume Next
    Set FindComponent = vbProj.VBComponents(componentName)
    On Error GoTo 0
End Function

'@label:source-project
'@sub-title The VBProject of the source workbook
'@return Object. The VBProject.
'@throws ProjectError.ErrorUnexpectedState When Excel refuses the read.
Private Function SourceProject() As Object
    On Error GoTo ProjectBlocked
    Set SourceProject = this.sourceBook.VBProject
    Exit Function

ProjectBlocked:
    ThrowError ProjectError.ErrorUnexpectedState, _
               TRUST_HINT & this.sourceBook.Name & TRUST_ACTION
End Function

'@label:target-project
'@sub-title The VBProject of the target workbook
'@return Object. The VBProject.
'@throws ProjectError.ErrorUnexpectedState When Excel refuses the read.
Private Function TargetProject() As Object
    On Error GoTo ProjectBlocked
    Set TargetProject = this.targetBook.VBProject
    Exit Function

ProjectBlocked:
    ThrowError ProjectError.ErrorUnexpectedState, _
               TRUST_HINT & this.targetBook.Name & TRUST_ACTION
End Function

'@label:read-code
'@sub-title Read a whole code module into one string
'@details
'One crossing for the whole module. Reading it line by line costs one crossing
'per line, so leave this as it is.
'@param codeMod Object. The CodeModule to read.
'@return String. Every line of the module, and an empty string for an empty one.
Private Function ReadCode(ByVal codeMod As Object) As String
    With codeMod
        If .CountOfLines > 0 Then ReadCode = .Lines(1, .CountOfLines)
    End With
End Function

'@label:replace-code
'@sub-title Write text into a code module and put the old text back on a failure
'@details
'DeleteLines 1, 0 on an empty module is what the line count guards against, and a
'workbook from Workbooks.Add carries an empty ThisWorkbook module. The text that
'is about to go is kept, so a raise between the delete and the write leaves the
'module holding what it held before.
'@param codeMod Object. The CodeModule to write into.
'@param codeContent String. The text to write.
Private Sub ReplaceCode(ByVal codeMod As Object, ByVal codeContent As String)
    Dim previousContent As String
    Dim failedNumber As Long
    Dim failedSource As String
    Dim failedDescription As String

    previousContent = ReadCode(codeMod)

    On Error GoTo RestorePrevious

    With codeMod
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If LenB(codeContent) > 0 Then .AddFromString codeContent
    End With
    Exit Sub

RestorePrevious:
    'The error is copied first, because the restore below runs its own
    'On Error Resume Next and that clears Err.
    failedNumber = Err.Number
    failedSource = Err.Source
    failedDescription = Err.Description

    On Error Resume Next
    With codeMod
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If LenB(previousContent) > 0 Then .AddFromString previousContent
    End With
    On Error GoTo 0

    Err.Raise failedNumber, failedSource, failedDescription
End Sub

'@label:log-info
'@sub-title Record a diagnostic entry, building the Checking object on first use
'@param label String. The message to record.
'@param scope Optional Byte. Severity from CheckingScope. Defaults to checkingNote.
Private Sub LogInfo(ByVal label As String, _
                    Optional ByVal scope As Byte = checkingNote)

    If this.internalChecks Is Nothing Then
        Set this.internalChecks = Checking.Create("Code transfer trace")
    End If

    this.internalChecks.Add CStr(this.checkCounter), label, scope
    this.checkCounter = this.checkCounter + 1
    this.hasChecks = True
End Sub

'@label:guard-not-sealed
'@sub-title Refuse a write to a creation-only binding after Create
'@param propName String. The binding the caller tried to write.
'@throws ProjectError.SomethingWentWrong When the instance is already sealed.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. The error code to raise.
'@param errorMessage String. Human-readable description of the failure.
Private Sub ThrowError(ByVal errNumber As Long, ByVal errorMessage As String)
    Err.Raise CLng(errNumber), CLASS_NAME, errorMessage
End Sub
