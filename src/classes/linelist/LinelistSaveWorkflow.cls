VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSaveWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Handles final save sequence for the Linelist workbook"

'@Folder("Linelist")
'@ModuleDescription("Handles final save sequence for the Linelist workbook")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistSaveWorkflow

'@section State and constants
'===============================================================================
Private Type TSaveWorkflow
    navigator As ILinelistSheetNavigator
    lifecycle As ILinelistLifecycleManager
End Type

Private this As TSaveWorkflow
Private Const CLASSNAME As String = "LinelistSaveWorkflow"


'@section Factory helpers
'===============================================================================
Public Function Create(ByVal navigator As ILinelistSheetNavigator, _
                       ByVal lifecycle As ILinelistLifecycleManager) As ILinelistSaveWorkflow

    Dim workflow As LinelistSaveWorkflow
    Set workflow = New LinelistSaveWorkflow
    workflow.Initialise navigator, lifecycle
    Set Create = workflow
End Function

Friend Sub Initialise(ByVal navigator As ILinelistSheetNavigator, _
                      ByVal lifecycle As ILinelistLifecycleManager)

    EnsureNavigator navigator
    EnsureLifecycle lifecycle

    Set this.navigator = navigator
    Set this.lifecycle = lifecycle
End Sub


'@section Core workflow
'===============================================================================
Private Sub ExecuteSave(ByVal context As ILinelistPreparationContext)
    Dim specs As ILinelistSpecs
    Dim mainObj As IMain
    Dim passwordProvider As IPasswords
    Dim workbook As Workbook
    Dim outputPath As String
    Dim password As String
    Dim saveResult As Boolean

    Set specs = EnsureSpecifications(context)
    Set mainObj = EnsureMainObject(specs)
    Set passwordProvider = EnsurePasswordProvider(specs)
    Set workbook = EnsureWorkbook(context)

    outputPath = BuildOutputPath(mainObj)
    password = mainObj.ValueOf("llpassword")

    ActivateSheets
    passwordProvider.ProtectWkb workbook
    saveResult = SaveWorkbook(workbook, outputPath, password)

    If Not saveResult Then Exit Sub

    this.lifecycle.Dispose False
End Sub

Private Sub ActivateSheets()
    On Error GoTo ActivationFail
        this.navigator.ActivateAdminSheet
        this.navigator.ActivateInstructionSheet
        On Error GoTo 0
        Exit Sub

ActivationFail:
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Failed to activate sheets prior to saving: " & Err.Description
End Sub

Private Function SaveWorkbook(ByVal workbook As Workbook, _
                              ByVal path As String, _
                              ByVal password As String) As Boolean
    On Error GoTo SaveFail
        workbook.SaveAs fileName:=path, _
                        FileFormat:=xlExcel12, _
                        Password:=password, _
                        ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        SaveWorkbook = True
        Exit Function

SaveFail:
    ThrowError ProjectError.ErrorUnexpectedState, _
              "Failed to save linelist workbook: " & Err.Description
End Function

Private Function BuildOutputPath(ByVal mainObj As IMain) As String
    Dim directoryPath As String
    Dim fileName As String

    directoryPath = mainObj.ValueOf("lldir")
    fileName = mainObj.ValueOf("llname")

    If LenB(Trim$(directoryPath)) = 0 Or LenB(Trim$(fileName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Linelist directory/name must be defined before saving"
    End If

    BuildOutputPath = directoryPath & Application.PathSeparator & fileName & ".xlsb"
End Function


'@section Dependency validation
'===============================================================================
Private Function EnsureSpecifications(ByVal context As ILinelistPreparationContext) As ILinelistSpecs
    Dim specs As ILinelistSpecs

    If context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Save context cannot be Nothing"
    End If

    Set specs = context.Specifications
    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Save context must expose specifications"
    End If

    Set EnsureSpecifications = specs
End Function

Private Function EnsureMainObject(ByVal specs As ILinelistSpecs) As IMain
    Dim mainObj As IMain

    Set mainObj = specs.MainObject
    If mainObj Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications must provide a main object"
    End If

    Set EnsureMainObject = mainObj
End Function

Private Function EnsurePasswordProvider(ByVal specs As ILinelistSpecs) As IPasswords
    Dim provider As IPasswords

    Set provider = specs.Password
    If provider Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications must provide a password provider"
    End If

    Set EnsurePasswordProvider = provider
End Function

Private Function EnsureWorkbook(ByVal context As ILinelistPreparationContext) As Workbook
    Dim linelist As ILinelist
    Dim accessor As ILinelistWorkbookAccessor
    Dim workbook As Workbook

    Set linelist = context.Linelist
    If linelist Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Context must expose a linelist reference"
    End If

    On Error GoTo MissingAccessor
        Set accessor = linelist
        On Error GoTo 0
        Set workbook = accessor.OutputWorkbook
        If workbook Is Nothing Then
            ThrowError ProjectError.ObjectNotInitialized, "Linelist output workbook cannot be Nothing"
        End If
        Set EnsureWorkbook = workbook
        Exit Function

MissingAccessor:
    Err.Clear
    ThrowError ProjectError.InvalidOperation, "Linelist does not expose ILinelistWorkbookAccessor"
End Function

Private Sub EnsureNavigator(ByVal navigator As ILinelistSheetNavigator)
    If navigator Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Sheet navigator cannot be Nothing"
    End If
End Sub

Private Sub EnsureLifecycle(ByVal lifecycle As ILinelistLifecycleManager)
    If lifecycle Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Lifecycle manager cannot be Nothing"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistSaveWorkflow_Save(ByVal context As ILinelistPreparationContext)
    ExecuteSave context
End Sub

