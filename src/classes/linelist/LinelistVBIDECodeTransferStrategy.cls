VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistVBIDECodeTransferStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Transfers VBA components using VBIDE import/export routines"

'@Folder("Linelist")
'@ModuleDescription("Transfers VBA components using VBIDE import/export routines")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed
'@depends ProjectError

Option Explicit

Implements ILinelistCodeTransferStrategy

Private Const CLASSNAME As String = "LinelistVBIDECodeTransferStrategy"
Private Const TEMPFILEPREFIX As String = "LLTransfer_"


'@section Factory helpers
'===============================================================================
Public Function Create() As ILinelistCodeTransferStrategy
    Dim strategy As LinelistVBIDECodeTransferStrategy
    Set strategy = New LinelistVBIDECodeTransferStrategy
    Set Create = strategy
End Function


'@section Strategy contract
'===============================================================================
Public Sub TransferComponent(ByVal componentName As String, _
                             ByVal scope As LinelistCodeScope, _
                             ByVal source As Workbook, _
                             ByVal target As Workbook)
    Dim extension As String
    Dim exportPath As String
    Dim sourceComponent As Object
    Dim targetComponent As Object
    Dim targetProject As Object

    extension = FileExtension(scope)
    exportPath = BuildTempPath(extension, source)

    Set sourceComponent = ResolveComponent(source, componentName)

    On Error GoTo CleanUp
        sourceComponent.Export exportPath
        Set targetProject = target.VBProject
        If ComponentExists(targetProject, componentName) Then
            targetProject.VBComponents.Remove targetProject.VBComponents(componentName)
        End If
        targetProject.VBComponents.Import exportPath
    On Error GoTo 0

CleanUp:
    SafeKill exportPath
    If Err.Number <> 0 Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Failed to transfer component '" & componentName & "' using VBIDE: " & Err.Description
    End If
End Sub

Public Sub TransferWorkbookModule(ByVal moduleName As String, _
                                  ByVal source As Workbook, _
                                  ByVal target As Workbook)
    Dim sourceModule As Object
    Dim targetModule As Object
    Dim moduleCode As String

    Set sourceModule = ResolveComponent(source, moduleName).CodeModule
    Set targetModule = ResolveComponent(target, moduleName).CodeModule

    moduleCode = sourceModule.Lines(1, sourceModule.CountOfLines)

    With targetModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        If Len(moduleCode) > 0 Then .AddFromString moduleCode
    End With
End Sub

Public Sub TransferForm(ByVal formName As String, _
                        ByVal source As Workbook, _
                        ByVal target As Workbook)
    Dim exportPath As String
    Dim sourceComponent As Object
    Dim targetProject As Object

    exportPath = BuildTempPath(".frm", source)
    Set sourceComponent = ResolveComponent(source, formName)

    On Error GoTo CleanUp
        sourceComponent.Export exportPath
        Set targetProject = target.VBProject
        If ComponentExists(targetProject, formName) Then
            targetProject.VBComponents.Remove targetProject.VBComponents(formName)
        End If
        targetProject.VBComponents.Import exportPath
    On Error GoTo 0

CleanUp:
    SafeKill exportPath
    SafeKill Replace$(exportPath, ".frm", ".frx")
    If Err.Number <> 0 Then
        ThrowError ProjectError.ErrorUnexpectedState, _
                  "Failed to transfer form '" & formName & "' using VBIDE: " & Err.Description
    End If
End Sub

Public Property Get Name() As String
    Name = "VBIDE"
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

Private Function ComponentExists(ByVal project As Object, ByVal componentName As String) As Boolean
    On Error GoTo Missing
        project.VBComponents componentName
        ComponentExists = True
        Exit Function
Missing:
    Err.Clear
    ComponentExists = False
End Function

Private Function TempDirectory(Optional ByVal referenceBook As Workbook) As String
    Dim candidate As String

    On Error Resume Next
        If Not referenceBook Is Nothing Then
            candidate = referenceBook.Path
        End If
    On Error GoTo 0

    If LenB(candidate) = 0 Then
        On Error Resume Next
            candidate = Application.Path
        On Error GoTo 0
    End If

    If LenB(candidate) = 0 Then
        candidate = CurDir$
    End If

    If LenB(candidate) = 0 Then
        ThrowError ProjectError.ErrorUnexpectedState, "Unable to determine a temporary directory for exports"
    End If

    If Right$(candidate, 1) <> Application.PathSeparator Then
        candidate = candidate & Application.PathSeparator
    End If

    TempDirectory = candidate
End Function

Private Function BuildTempPath(ByVal extension As String, _
                               Optional ByVal referenceBook As Workbook) As String
    Dim folder As String
    Dim fileName As String

    Randomize
    folder = TempDirectory(referenceBook)
    fileName = TEMPFILEPREFIX & Format$(Now, "yyyymmdd_hhnnss") & "_" & CLng(Rnd * 100000)
    BuildTempPath = folder & fileName & extension
End Function

Private Function FileExtension(ByVal scope As LinelistCodeScope) As String
    Select Case scope
        Case codeScopeClass
            FileExtension = ".cls"
        Case codeScopeModule
            FileExtension = ".bas"
        Case Else
            ThrowError ProjectError.InvalidArgument, "Unsupported code scope value " & CStr(scope)
    End Select
End Function

Private Sub SafeKill(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub
    On Error Resume Next
        Kill path
    On Error GoTo 0
End Sub

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
