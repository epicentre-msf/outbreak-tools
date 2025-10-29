VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalysisTableSpecBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis")
'@ModuleDescription("Creates ITablesSpecs instances for analysis traversal leveraging legacy TablesSpecs factory")
'@depends ProjectError
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITableSpecsBuilder

Private Const CLASSNAME As String = "AnalysisTableSpecBuilder"

'@section Factory
'===============================================================================
'Return a builder ready to create specifications using the legacy TablesSpecs class.
Public Function Create() As ITableSpecsBuilder
    With New AnalysisTableSpecBuilder
        Set Create = .Self
    End With
End Function

'Expose the current instance through the builder interface.
Public Property Get Self() As ITableSpecsBuilder
    Set Self = Me
End Property

'@section Validation helpers
'===============================================================================
'Ensure a non-empty range is provided for the header row.
Private Sub ValidateRange(ByVal candidate As Range, ByVal argumentName As String)
    If candidate Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If

    If StrComp(TypeName(candidate), "Range", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " must be a Range instance"
    End If
End Sub

'Ensure the linelist reference is valid.
Private Sub ValidateLinelist(ByVal linelistSpecs As ILinelistSpecs)
    If linelistSpecs Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), "linelistSpecs cannot be Nothing"
    End If

    If LenB(Trim$(TypeName(linelistSpecs))) = 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), _
                   "linelistSpecs must expose a concrete implementation"
    End If
End Sub

'Centralised helper for raising consistent project errors.
Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASSNAME, messageText
End Sub

'@section Builder implementation
'===============================================================================
Private Function ITableSpecsBuilder_CreateSpec(ByVal headerRange As Range, _
                                               ByVal tableRange As Range, _
                                               ByVal linelistSpecs As ILinelistSpecs) As ITablesSpecs
    ValidateRange headerRange, "headerRange"
    ValidateRange tableRange, "tableRange"
    ValidateLinelist linelistSpecs

    On Error GoTo CleanFail
        Set ITableSpecsBuilder_CreateSpec = TablesSpecs.Create(headerRange, tableRange, linelistSpecs)
    On Error GoTo 0

    If ITableSpecsBuilder_CreateSpec Is Nothing Then
        RaiseError CLng(ProjectError.ErrorUnexpectedState), _
                   "TablesSpecs.Create returned Nothing for row " & CStr(tableRange.Row)
    End If

    Exit Function

CleanFail:
    RaiseError CLng(ProjectError.ErrorUnexpectedState), _
               "Failed to create TableSpecs for row " & CStr(tableRange.Row)
End Function
