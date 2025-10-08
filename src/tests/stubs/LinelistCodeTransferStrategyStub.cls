VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistCodeTransferStrategyStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Records calls made through ILinelistCodeTransferStrategy for testing")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, HungarianNotation

Option Explicit

Implements ILinelistCodeTransferStrategy

Private Type TState
    componentCalls As BetterArray
    workbookModuleCalls As BetterArray
    formCalls As BetterArray
    strategyName As String
End Type

Private this As TState

Public Sub Initialise(ByVal name As String)
    Set this.componentCalls = New BetterArray
    this.componentCalls.LowerBound = 1

    Set this.workbookModuleCalls = New BetterArray
    this.workbookModuleCalls.LowerBound = 1

    Set this.formCalls = New BetterArray
    this.formCalls.LowerBound = 1

    this.strategyName = name
End Sub

Public Function ComponentLog() As BetterArray
    Set ComponentLog = this.componentCalls
End Function

Public Function WorkbookModuleLog() As BetterArray
    Set WorkbookModuleLog = this.workbookModuleCalls
End Function

Public Function FormLog() As BetterArray
    Set FormLog = this.formCalls
End Function

Private Sub RecordCall(ByVal log As BetterArray, ByVal value As String)
    log.Push value
End Sub

Private Sub ILinelistCodeTransferStrategy_TransferComponent(ByVal componentName As String, _
                                                            ByVal scope As LinelistCodeScope, _
                                                            ByVal source As Workbook, _
                                                            ByVal target As Workbook)
    RecordCall this.componentCalls, componentName & "|" & CStr(scope)
End Sub

Private Sub ILinelistCodeTransferStrategy_TransferWorkbookModule(ByVal moduleName As String, _
                                                                 ByVal source As Workbook, _
                                                                 ByVal target As Workbook)
    RecordCall this.workbookModuleCalls, moduleName
End Sub

Private Sub ILinelistCodeTransferStrategy_TransferForm(ByVal formName As String, _
                                                       ByVal source As Workbook, _
                                                       ByVal target As Workbook)
    RecordCall this.formCalls, formName
End Sub

Private Property Get ILinelistCodeTransferStrategy_Name() As String
    ILinelistCodeTransferStrategy_Name = this.strategyName
End Property

