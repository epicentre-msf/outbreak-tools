VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Collects lightweight log entries for disease workflows"

'@Folder("Msetup")
'@ModuleDescription("Collects lightweight log entries for disease workflows")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseLogger

'@section Types
'===============================================================================

Private Type TLogEntry
    Timestamp As Date
    Operation As String
    Severity As DiseaseLogSeverity
    Message As String
    Details As BetterArray
End Type

Private Type TLogger
    Entries As BetterArray
End Type

Private this As TLogger

'@section Lifecycle
'===============================================================================

Private Sub Class_Initialize()
    ResetEntries
End Sub

'@section Logging
'===============================================================================

Public Sub Record(ByVal operation As String, _
                  ByVal severity As DiseaseLogSeverity, _
                  ByVal message As String, _
                  Optional ByVal details As BetterArray)

    Dim entry As TLogEntry

    ValidateOperation operation
    ValidateMessage message

    entry.Timestamp = Now
    entry.Operation = operation
    entry.Severity = severity
    entry.Message = message
    Set entry.Details = CloneDetails(details)

    AppendEntry entry
End Sub

'@section Query
'===============================================================================

Public Property Get Entries() As BetterArray
    Entries = CloneEntries()
End Property

Public Property Get HasEntries() As Boolean
    If this.Entries Is Nothing Then Exit Property
    HasEntries = (this.Entries.Length > 0)
End Property

'@section Maintenance
'===============================================================================

Public Sub Clear()
    ResetEntries
End Sub

'@section Internal Helpers
'===============================================================================

Private Sub AppendEntry(ByVal entry As TLogEntry)
    Dim snapshot As BetterArray
    Set snapshot = New BetterArray

    snapshot.Push entry.Timestamp
    snapshot.Push entry.Operation
    snapshot.Push entry.Severity
    snapshot.Push entry.Message
    snapshot.Push entry.Details.Clone()

    this.Entries.Push snapshot
End Sub

Private Function CloneEntries() As BetterArray
    Dim cloned As BetterArray
    Dim index As Long
    Dim item As BetterArray

    Set cloned = New BetterArray

    If this.Entries Is Nothing Then
        Set CloneEntries = cloned
        Exit Function
    End If

    For index = this.Entries.LowerBound To this.Entries.UpperBound
        Set item = this.Entries.Item(index)
        cloned.Push item.Clone()
    Next index

    Set CloneEntries = cloned
End Function

Private Function CloneDetails(ByVal details As BetterArray) As BetterArray
    If details Is Nothing Then
        Set CloneDetails = New BetterArray
    Else
        Set CloneDetails = details.Clone()
    End If
End Function

Private Sub ResetEntries()
    Set this.Entries = New BetterArray
    this.Entries.lowerBound = 0
End Sub

Private Sub ValidateOperation(ByVal operation As String)
    If LenB(operation) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Operation name cannot be empty"
    End If
End Sub

Private Sub ValidateMessage(ByVal message As String)
    If LenB(message) = 0 Then
        ThrowError ProjectError.InvalidArgument, "Log message cannot be empty"
    End If
End Sub

'@section ErrorHandling
'===============================================================================

Private Sub ThrowError(ByVal errNumber As ProjectError, ByVal errMessage As String)
    Err.Raise errNumber, "DiseaseLogger", errMessage
End Sub

'@section Interface
'===============================================================================

Private Sub IDiseaseLogger_Record(ByVal operation As String, ByVal severity As DiseaseLogSeverity, ByVal message As String, Optional ByVal details As BetterArray)
    Record operation, severity, message, details
End Sub

Private Property Get IDiseaseLogger_Entries() As BetterArray
    Set IDiseaseLogger_Entries = Entries
End Property

Private Property Get IDiseaseLogger_HasEntries() As Boolean
    IDiseaseLogger_HasEntries = HasEntries
End Property

Private Sub IDiseaseLogger_Clear()
    Clear
End Sub
