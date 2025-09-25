VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeSeriesTableSpecsFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Analysis.Graphs")
'@ModuleDescription("Creates ITablesSpecs instances for time-series graph rows")
'@depends ProjectError
'@IgnoreModule SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Option Explicit

Implements ITimeSeriesTableSpecsFactory

Private Const CLASS_NAME As String = "TimeSeriesTableSpecsFactory"

'@section Factory
'===============================================================================
Public Function CreateInstance() As ITimeSeriesTableSpecsFactory
    With New TimeSeriesTableSpecsFactory
        Set CreateInstance = .Self
    End With
End Function

Public Property Get Self() As ITimeSeriesTableSpecsFactory
    Set Self = Me
End Property

'@section Validation helpers
'===============================================================================
Private Sub ValidateRange(ByVal target As Range, ByVal argumentName As String)
    If target Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub ValidateListRow(ByVal listRow As ListRow, ByVal argumentName As String)
    If listRow Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
        Exit Sub
    End If

    If StrComp(TypeName(listRow), "ListRow", vbBinaryCompare) <> 0 Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " must be a ListRow"
    End If
End Sub

Private Sub ValidateLinelist(ByVal linelist As ILinelistSpecs, ByVal argumentName As String)
    If linelist Is Nothing Then
        RaiseError CLng(ProjectError.InvalidArgument), argumentName & " cannot be Nothing"
    End If
End Sub

Private Sub RaiseError(ByVal errNumber As Long, ByVal messageText As String)
    Err.Raise errNumber, CLASS_NAME, messageText
End Sub

'@section Interface implementation
'===============================================================================
Private Function ITimeSeriesTableSpecsFactory_Create(ByVal headerRange As Range, _
                                                     ByVal listRow As ListRow, _
                                                     ByVal linelist As ILinelistSpecs) As ITablesSpecs
    Dim tableRange As Range

    ValidateRange headerRange, "headerRange"
    ValidateListRow listRow, "listRow"
    ValidateLinelist linelist, "linelist"

    Set tableRange = listRow.Range
    Set ITimeSeriesTableSpecsFactory_Create = TablesSpecs.Create(headerRange, tableRange, linelist)
End Function

Private Property Get ITimeSeriesTableSpecsFactory_Self() As ITimeSeriesTableSpecsFactory
    Set ITimeSeriesTableSpecsFactory_Self = Self
End Property
