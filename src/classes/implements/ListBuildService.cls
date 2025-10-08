VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListBuildService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Linelist.Shared")
'@ModuleDescription("High-level service composing the list builder factory and coordinator")
'@depends ProjectError, ListBuilderFactory
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Private Type TServiceState
    factory As ListBuilderFactory
End Type

Private this As TServiceState
Private Const CLASSNAME As String = "ListBuildService"

'@section Factory
'===============================================================================
'Create a service optionally using the supplied factory.
Public Function Create(Optional ByVal factory As ListBuilderFactory = Nothing) As ListBuildService
    With New ListBuildService
        If factory Is Nothing Then
            Set .Factory = ListBuilderFactory.Create
        Else
            Set .Factory = factory
        End If
        Set Create = .Self
    End With
End Function

'Expose the instance for fluent usage.
Public Property Get Self() As ListBuildService
    Set Self = Me
End Property

'Assign the factory after validation.
Public Property Set Factory(ByVal factory As ListBuilderFactory)
    If factory Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "factory cannot be Nothing"
    End If
    Set this.factory = factory
End Property

'@section Build orchestration
'===============================================================================
'Build the specified worksheet and return the processed sections.
Public Function Build(ByVal linelist As ILinelist, _
                      ByVal sheetName As String) As BetterArray

    ValidateLinelist linelist
    ValidateWorksheetName sheetName

    Dim context As IListBuildContext
    Set context = this.factory.CreateContext(linelist, sheetName)

    Dim coordinator As IListBuilderCoordinator
    Set coordinator = this.factory.CreateCoordinator(context)

    Set Build = coordinator.Build(context)
End Function

'@section Helpers
'===============================================================================
Private Sub ValidateLinelist(ByVal linelist As ILinelist)
    If linelist Is Nothing Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "linelist cannot be Nothing"
    End If
End Sub

Private Sub ValidateWorksheetName(ByVal sheetName As String)
    If Trim$(sheetName) = vbNullString Then
        ThrowError CLng(ProjectError.InvalidArgument), _
                  "sheetName cannot be empty"
    End If
End Sub

Private Sub ThrowError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

