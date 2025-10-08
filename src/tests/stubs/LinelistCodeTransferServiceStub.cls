VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistCodeTransferServiceStub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Stub ILinelistCodeTransferService capturing transfer requests for verification")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ILinelistCodeTransferService

Private Type TService
    componentCalls As BetterArray
    workbookModuleCalls As BetterArray
    formCalls As BetterArray
    strategyName As String
End Type

Private this As TService

Public Sub Initialise(Optional ByVal name As String = "StubStrategy")
    Set this.componentCalls = New BetterArray
    this.componentCalls.LowerBound = 1

    Set this.workbookModuleCalls = New BetterArray
    this.workbookModuleCalls.LowerBound = 1

    Set this.formCalls = New BetterArray
    this.formCalls.LowerBound = 1

    this.strategyName = name
End Sub

Public Function ComponentCalls() As BetterArray
    Set ComponentCalls = this.componentCalls
End Function

Public Function FormCalls() As BetterArray
    Set FormCalls = this.formCalls
End Function

Public Function WorkbookModuleCalls() As BetterArray
    Set WorkbookModuleCalls = this.workbookModuleCalls
End Function

Private Sub RecordComponents(ByVal componentNames As BetterArray, ByVal scope As LinelistCodeScope)
    Dim index As Long

    For index = componentNames.LowerBound To componentNames.UpperBound
        this.componentCalls.Push CStr(componentNames.Item(index)) & "|" & CStr(scope)
    Next index
End Sub

Private Sub ILinelistCodeTransferService_TransferComponent(ByVal componentName As String, ByVal scope As LinelistCodeScope, ByVal source As Workbook, ByVal target As Workbook)
    Dim singleItem As BetterArray
    Set singleItem = New BetterArray
    singleItem.LowerBound = 1
    singleItem.Push componentName
    RecordComponents singleItem, scope
End Sub

Private Sub ILinelistCodeTransferService_TransferWorkbookModule(ByVal moduleName As String, ByVal source As Workbook, ByVal target As Workbook)
    this.workbookModuleCalls.Push moduleName
End Sub

Private Sub ILinelistCodeTransferService_TransferForm(ByVal formName As String, ByVal source As Workbook, ByVal target As Workbook)
    this.formCalls.Push formName
End Sub

Private Sub ILinelistCodeTransferService_TransferComponents(ByVal componentNames As BetterArray, ByVal scope As LinelistCodeScope, ByVal source As Workbook, ByVal target As Workbook)
    If Not componentNames Is Nothing Then
        RecordComponents componentNames, scope
    End If
End Sub

Private Property Get ILinelistCodeTransferService_StrategyName() As String
    ILinelistCodeTransferService_StrategyName = this.strategyName
End Property

