VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDiseaseLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface describing lightweight logging for disease workflows"

'@Interface
'@Folder("Msetup")
'@ModuleDescription("Interface describing lightweight logging for disease workflows")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit

'@section Enumerations
'===============================================================================

Public Enum DiseaseLogSeverity
    DiseaseLogInfo = 1
    DiseaseLogWarning = 2
    DiseaseLogError = 3
End Enum

'@section Logging
'===============================================================================

'@sub-title Record a new log entry.
'@param operation String operation identifier (e.g., "Import", "Export").
'@param severity DiseaseLogSeverity describing the entry importance.
'@param message String descriptive message.
'@param details Optional BetterArray containing additional context strings.
Public Sub Record(ByVal operation As String, _
                  ByVal severity As DiseaseLogSeverity, _
                  ByVal message As String, _
                  Optional ByVal details As BetterArray)
End Sub

'@section Query
'===============================================================================

'@sub-title Retrieve recorded entries.
'@return BetterArray clone containing log entry snapshots.
Public Property Get Entries() As BetterArray
End Property

'@sub-title Indicates whether any entries have been recorded.
Public Property Get HasEntries() As Boolean
End Property

'@section Maintenance
'===============================================================================

'@sub-title Clear all recorded entries.
Public Sub Clear()
End Sub
