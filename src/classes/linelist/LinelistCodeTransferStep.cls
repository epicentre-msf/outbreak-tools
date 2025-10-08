VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistCodeTransferStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Preparation step that synchronises VBA components into the output linelist workbook"

'@Folder("Linelist")
'@ModuleDescription("Preparation step that synchronises VBA components into the output linelist workbook")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistPreparationStep

'@section State and constants
'===============================================================================
Private Type TStep
    service As ILinelistCodeTransferService
    plan As ILinelistCodeTransferPlan
End Type

Private this As TStep
Private Const CLASSNAME As String = "LinelistCodeTransferStep"


'@section Factory helpers
'===============================================================================
Public Function Create(ByVal service As ILinelistCodeTransferService, _
                       Optional ByVal plan As ILinelistCodeTransferPlan) As ILinelistPreparationStep

    Dim stepObject As LinelistCodeTransferStep
    Set stepObject = New LinelistCodeTransferStep
    stepObject.Initialise service, plan
    Set Create = stepObject
End Function

Friend Sub Initialise(ByVal service As ILinelistCodeTransferService, _
                      ByVal plan As ILinelistCodeTransferPlan)

    EnsureService service

    If plan Is Nothing Then
        Set this.plan = LinelistDefaultCodeTransferPlan.Create
    Else
        Set this.plan = plan
    End If

    Set this.service = service
End Sub


'@section Execution helpers
'===============================================================================
Private Sub ExecuteTransfer(ByVal context As ILinelistPreparationContext)
    Dim sourceWorkbook As Workbook
    Dim targetWorkbook As Workbook

    Set sourceWorkbook = ResolveSourceWorkbook(context)
    Set targetWorkbook = ResolveTargetWorkbook(context)

    TransferClassModules sourceWorkbook, targetWorkbook
    TransferStandardModules sourceWorkbook, targetWorkbook
    TransferForms sourceWorkbook, targetWorkbook
    TransferWorkbookModule sourceWorkbook, targetWorkbook
End Sub

Private Sub TransferClassModules(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook)
    Dim components As BetterArray

    Set components = this.plan.ClassComponentNames
    If components.Length = 0 Then Exit Sub

    this.service.TransferComponents components, codeScopeClass, sourceWorkbook, targetWorkbook
End Sub

Private Sub TransferStandardModules(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook)
    Dim modules As BetterArray

    Set modules = this.plan.ModuleComponentNames
    If modules.Length = 0 Then Exit Sub

    this.service.TransferComponents modules, codeScopeModule, sourceWorkbook, targetWorkbook
End Sub

Private Sub TransferForms(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook)
    Dim forms As BetterArray
    Dim index As Long

    Set forms = this.plan.FormNames
    If forms.Length = 0 Then Exit Sub

    For index = forms.LowerBound To forms.UpperBound
        this.service.TransferForm CStr(forms.Item(index)), sourceWorkbook, targetWorkbook
    Next index
End Sub

Private Sub TransferWorkbookModule(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook)
    Dim moduleName As String

    moduleName = Trim$(this.plan.WorkbookModuleName)
    If LenB(moduleName) = 0 Then Exit Sub

    this.service.TransferWorkbookModule moduleName, sourceWorkbook, targetWorkbook
End Sub


'@section Resolution helpers
'===============================================================================
Private Function ResolveSourceWorkbook(ByVal context As ILinelistPreparationContext) As Workbook
    Dim specs As ILinelistSpecs
    Dim workbook As Workbook

    Set specs = context.Specifications
    If specs Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Preparation context does not expose specifications"
    End If

    Set workbook = specs.Workbook
    If workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Specifications workbook reference cannot be Nothing"
    End If

    Set ResolveSourceWorkbook = workbook
End Function

Private Function ResolveTargetWorkbook(ByVal context As ILinelistPreparationContext) As Workbook
    Dim linelist As ILinelist
    Dim accessor As ILinelistWorkbookAccessor
    Dim workbook As Workbook

    Set linelist = context.Linelist
    If linelist Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Preparation context must expose a linelist reference"
    End If

    Set accessor = CastWorkbookAccessor(linelist)
    Set workbook = accessor.OutputWorkbook

    If workbook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Linelist output workbook cannot be Nothing"
    End If

    Set ResolveTargetWorkbook = workbook
End Function

Private Function CastWorkbookAccessor(ByVal candidate As ILinelist) As ILinelistWorkbookAccessor
    Dim accessor As ILinelistWorkbookAccessor

    On Error GoTo MissingInterface
        Set accessor = candidate
        On Error GoTo 0
        Set CastWorkbookAccessor = accessor
        Exit Function

MissingInterface:
    Err.Clear
    ThrowError ProjectError.InvalidOperation, _
              "Linelist instance " & TypeName(candidate) & " does not expose ILinelistWorkbookAccessor"
End Function


'@section Validation helpers
'===============================================================================
Private Sub EnsureService(ByVal service As ILinelistCodeTransferService)
    If service Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Code transfer service cannot be Nothing"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistPreparationStep_Execute(ByVal context As ILinelistPreparationContext)
    If context Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Preparation context cannot be Nothing"
    End If

    ExecuteTransfer context
End Sub
