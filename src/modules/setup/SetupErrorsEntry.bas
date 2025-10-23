Attribute VB_Name = "SetupErrorsEntry"
Attribute VB_Description = "Entry points delegating setup consistency checks to SetupErrors class"

Option Explicit

'@Folder("Setup")
'@ModuleDescription("Entry points delegating setup consistency checks to SetupErrors class")
'@depends SetupErrors, ISetupErrors, ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, UseMeaningfulName

'@section Factory helpers
'===============================================================================
'@sub-title Resolve the workbook that will be analysed.
'@param hostBook Optional workbook reference. Defaults to ThisWorkbook.
'@return Workbook reference guaranteed to be non-Nothing.
Private Function ResolveWorkbook(Optional ByVal hostBook As Workbook) As Workbook
    If hostBook Is Nothing Then
        Set hostBook = ThisWorkbook
    End If

    If hostBook Is Nothing Then
        ThrowError ProjectError.ObjectNotInitialized, "Host workbook reference is required"
    End If

    Set ResolveWorkbook = hostBook
End Function

'@sub-title Instantiate a SetupErrors checker.
'@param hostBook Workbook containing setup sheets to inspect.
'@return ISetupErrors instance ready to execute.
Private Function CreateChecker(Optional ByVal hostBook As Workbook) As ISetupErrors
    Dim checker As ISetupErrors

    Set checker = SetupErrors.Create(ResolveWorkbook(hostBook))
    Set CreateChecker = checker
End Function

'@section Public API
'===============================================================================
'@sub-title Backwards-compatible entry point matching the legacy module signature.
Public Sub CheckTheSetup()
    RunSetupChecks ThisWorkbook
End Sub

'@sub-title Execute setup checks against the provided workbook.
'@param hostBook Optional workbook. When omitted, ThisWorkbook is used.
Public Sub RunSetupChecks(Optional ByVal hostBook As Workbook)
    Dim checker As ISetupErrors
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    On Error GoTo RunFailed
        Set checker = CreateChecker(hostBook)
        checker.Run
    Exit Sub

RunFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    If errNumber <> 0 Then
        Err.Raise errNumber, errSource, errDescription
    End If
End Sub

