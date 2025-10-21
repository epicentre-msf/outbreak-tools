VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisErrorHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Utility helper centralising common ProjectError raises for Analysis components")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Const CLASSNAME As String = "AnalysisErrorHelper"

'@section Invalid argument helpers
'===============================================================================
Public Sub RaiseInvalidArgument(ByVal argumentName As String)
    Err.Raise CLng(ProjectError.InvalidArgument), CLASSNAME, Trim$(argumentName) & " is invalid for this operation"
End Sub

Public Sub RaiseMissingDependency(ByVal dependencyName As String)
    Err.Raise CLng(ProjectError.ObjectNotInitialized), CLASSNAME, Trim$(dependencyName) & " must be provided before invoking this operation"
End Sub

'@section State helpers
'===============================================================================
Public Sub RaiseUnexpectedState(ByVal messageText As String)
    Err.Raise CLng(ProjectError.ErrorUnexpectedState), CLASSNAME, messageText
End Sub
