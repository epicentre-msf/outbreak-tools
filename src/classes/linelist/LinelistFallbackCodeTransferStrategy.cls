VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistFallbackCodeTransferStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Fallback strategy that reuses existing workbook components without VBIDE import/export"

'@Folder("Linelist")
'@ModuleDescription("Fallback strategy that reuses existing workbook components without VBIDE import/export")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistCodeTransferStrategy

Private Const CLASSNAME As String = "LinelistFallbackCodeTransferStrategy"


'@section Factory helpers
'===============================================================================
Public Function Create() As ILinelistCodeTransferStrategy
    Dim strategy As LinelistFallbackCodeTransferStrategy
    Set strategy = New LinelistFallbackCodeTransferStrategy
    Set Create = strategy
End Function


'@section Strategy contract
'===============================================================================
Public Sub TransferComponent(ByVal componentName As String, _
                             ByVal scope As LinelistCodeScope, _
                             ByVal source As Workbook, _
                             ByVal target As Workbook)
    Dim sourceComponent As Object
    Dim targetComponent As Object
    Dim codeModule As Object
    Dim moduleText As String

    Set sourceComponent = ResolveComponent(source, componentName)
    Set targetComponent = EnsureTargetComponent(target, componentName)

    Set codeModule = sourceComponent.CodeModule
    moduleText = codeModule.Lines(1, codeModule.CountOfLines)

    With targetComponent.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(moduleText) > 0 Then .AddFromString moduleText
    End With
End Sub

Public Sub TransferWorkbookModule(ByVal moduleName As String, _
                                  ByVal source As Workbook, _
                                  ByVal target As Workbook)
    Dim sourceModule As Object
    Dim targetModule As Object
    Dim moduleText As String

    Set sourceModule = ResolveComponent(source, moduleName).CodeModule
    Set targetModule = ResolveComponent(target, moduleName).CodeModule

    moduleText = sourceModule.Lines(1, sourceModule.CountOfLines)

    With targetModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(moduleText) > 0 Then .AddFromString moduleText
    End With
End Sub

Public Sub TransferForm(ByVal formName As String, _
                        ByVal source As Workbook, _
                        ByVal target As Workbook)
    'Fallback strategy assumes forms already exist in the target workbook.
    If Not ComponentExists(target, formName) Then
        ThrowError ProjectError.ElementNotFound, _
                  "Form '" & formName & "' must be preloaded in the target workbook when VBIDE automation is unavailable"
    End If
End Sub

Public Property Get Name() As String
    Name = "Fallback"
End Property


'@section Helpers
'===============================================================================
Private Function ResolveComponent(ByVal workbook As Workbook, ByVal componentName As String) As Object
    Dim vbProject As Object

    On Error GoTo Missing
        Set vbProject = workbook.VBProject
        Set ResolveComponent = vbProject.VBComponents(componentName)
    On Error GoTo 0
    Exit Function

Missing:
    ThrowError ProjectError.ElementNotFound, _
              "Component '" & componentName & "' not found in workbook '" & workbook.Name & "'"
End Function

Private Function EnsureTargetComponent(ByVal workbook As Workbook, ByVal componentName As String) As Object
    If ComponentExists(workbook, componentName) Then
        Set EnsureTargetComponent = workbook.VBProject.VBComponents(componentName)
    Else
        ThrowError ProjectError.ElementNotFound, _
                  "Target workbook must contain component '" & componentName & "' when VBIDE automation is unavailable"
    End If
End Function

Private Function ComponentExists(ByVal workbook As Workbook, ByVal componentName As String) As Boolean
    On Error GoTo Missing
        workbook.VBProject.VBComponents componentName
        ComponentExists = True
        Exit Function
Missing:
    Err.Clear
    ComponentExists = False
End Function

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal message As String)
    Err.Raise errNumber, CLASSNAME, message
End Sub


'@section Interface implementation
'===============================================================================
Private Sub ILinelistCodeTransferStrategy_TransferComponent(ByVal componentName As String, _
                                                            ByVal scope As LinelistCodeScope, _
                                                            ByVal source As Workbook, _
                                                            ByVal target As Workbook)
    TransferComponent componentName, scope, source, target
End Sub

Private Sub ILinelistCodeTransferStrategy_TransferWorkbookModule(ByVal moduleName As String, _
                                                                 ByVal source As Workbook, _
                                                                 ByVal target As Workbook)
    TransferWorkbookModule moduleName, source, target
End Sub

Private Sub ILinelistCodeTransferStrategy_TransferForm(ByVal formName As String, _
                                                       ByVal source As Workbook, _
                                                       ByVal target As Workbook)
    TransferForm formName, source, target
End Sub

Private Property Get ILinelistCodeTransferStrategy_Name() As String
    ILinelistCodeTransferStrategy_Name = Name
End Property
