VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistCodeTransferService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Coordinates VBA component transfers using the appropriate strategy for the host"

'@Folder("Linelist")
'@ModuleDescription("Coordinates VBA component transfers using the appropriate strategy for the host")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistCodeTransferService

'@section State and constants
'===============================================================================
Private Type TCodeTransferService
    strategy As ILinelistCodeTransferStrategy
End Type

Private this As TCodeTransferService
Private Const CLASSNAME As String = "LinelistCodeTransferService"


'@section Factory helpers
'===============================================================================
Public Function Create(Optional ByVal strategy As ILinelistCodeTransferStrategy) As ILinelistCodeTransferService
    Dim service As LinelistCodeTransferService
    Set service = New LinelistCodeTransferService
    service.Initialise strategy
    Set Create = service
End Function

Friend Sub Initialise(Optional ByVal strategy As ILinelistCodeTransferStrategy)
    If strategy Is Nothing Then
        Set this.strategy = DetermineStrategy
    Else
        Set this.strategy = strategy
    End If
End Sub


'@section Public API
'===============================================================================
Public Sub TransferComponent(ByVal componentName As String, _
                             ByVal scope As LinelistCodeScope, _
                             ByVal source As Workbook, _
                             ByVal target As Workbook)
    ValidateArguments componentName, source, target
    this.strategy.TransferComponent componentName, scope, source, target
End Sub

Public Sub TransferWorkbookModule(ByVal moduleName As String, _
                                  ByVal source As Workbook, _
                                  ByVal target As Workbook)
    ValidateArguments moduleName, source, target
    this.strategy.TransferWorkbookModule moduleName, source, target
End Sub

Public Sub TransferForm(ByVal formName As String, _
                        ByVal source As Workbook, _
                        ByVal target As Workbook)
    ValidateArguments formName, source, target
    this.strategy.TransferForm formName, source, target
End Sub

Public Sub TransferComponents(ByVal componentNames As BetterArray, _
                              ByVal scope As LinelistCodeScope, _
                              ByVal source As Workbook, _
                              ByVal target As Workbook)
    Dim index As Long

    If componentNames Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Component list cannot be Nothing"
    End If

    For index = componentNames.LowerBound To componentNames.UpperBound
        TransferComponent CStr(componentNames.Item(index)), scope, source, target
    Next index
End Sub

Public Property Get StrategyName() As String
    If this.strategy Is Nothing Then
        StrategyName = "(none)"
    Else
        StrategyName = this.strategy.Name
    End If
End Property


'@section Internal helpers
'===============================================================================
Private Sub ValidateArguments(ByVal componentName As String, _
                              ByVal source As Workbook, _
                              ByVal target As Workbook)
    If LenB(Trim$(componentName)) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Component name cannot be empty"
    End If

    EnsureWorkbook source, "Source workbook cannot be Nothing"
    EnsureWorkbook target, "Target workbook cannot be Nothing"
End Sub

Private Sub EnsureWorkbook(ByVal candidate As Workbook, ByVal message As String)
    If candidate Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, message
    End If

    If StrComp(TypeName(candidate), "Workbook", vbTextCompare) <> 0 Then
        ThrowError ProjectError.InvalidArgument, "Workbook argument must be an Excel workbook"
    End If
End Sub

Private Function DetermineStrategy() As ILinelistCodeTransferStrategy
    If CanAccessVBIDE() Then
        Set DetermineStrategy = LinelistVBIDECodeTransferStrategy.Create
    Else
        Set DetermineStrategy = LinelistFallbackCodeTransferStrategy.Create
    End If
End Function

Private Function CanAccessVBIDE() As Boolean
    On Error GoTo NoAccess
        Dim vbeObject As Object
        Set vbeObject = Application.VBE
        CanAccessVBIDE = Not vbeObject Is Nothing
    On Error GoTo 0
    Exit Function
NoAccess:
    Err.Clear
    CanAccessVBIDE = False
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistCodeTransferService_TransferComponent(ByVal componentName As String, _
                                                           ByVal scope As LinelistCodeScope, _
                                                           ByVal source As Workbook, _
                                                           ByVal target As Workbook)
    TransferComponent componentName, scope, source, target
End Sub

Private Sub ILinelistCodeTransferService_TransferWorkbookModule(ByVal moduleName As String, _
                                                                ByVal source As Workbook, _
                                                                ByVal target As Workbook)
    TransferWorkbookModule moduleName, source, target
End Sub

Private Sub ILinelistCodeTransferService_TransferForm(ByVal formName As String, _
                                                      ByVal source As Workbook, _
                                                      ByVal target As Workbook)
    TransferForm formName, source, target
End Sub

Private Sub ILinelistCodeTransferService_TransferComponents(ByVal componentNames As BetterArray, _
                                                            ByVal scope As LinelistCodeScope, _
                                                            ByVal source As Workbook, _
                                                            ByVal target As Workbook)
    TransferComponents componentNames, scope, source, target
End Sub

Private Property Get ILinelistCodeTransferService_StrategyName() As String
    ILinelistCodeTransferService_StrategyName = StrategyName
End Property
