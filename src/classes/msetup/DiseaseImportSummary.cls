VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiseaseImportSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Collects the outcome of disease import merge operations"

'@Folder("Msetup")
'@ModuleDescription("Collects the outcome of disease import merge operations")
'@IgnoreModule UnrecognizedAnnotation, ProcedureNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Option Explicit
Implements IDiseaseImportSummary

'@section State
'===============================================================================

Private Type TImportSummary
    Missing As BetterArray
    Updated As BetterArray
    Appended As BetterArray
    RequiresReport As Boolean
End Type

Private this As TImportSummary

'@section Factory
'===============================================================================

'@sub-title Return a new summary instance as its interface.
Public Function Create() As IDiseaseImportSummary
    Dim summary As DiseaseImportSummary
    Set summary = New DiseaseImportSummary
    summary.Initialise
    Set Create = summary.Self
End Function

'@section Public API
'===============================================================================

'@sub-title Provide the interface reference to the current instance.
Public Property Get Self() As IDiseaseImportSummary
    Set Self = Me
End Property

'@sub-title Ensure arrays are initialised before use.
Public Sub Initialise()
    If this.Missing Is Nothing Then Set this.Missing = New BetterArray
    If this.Updated Is Nothing Then Set this.Updated = New BetterArray
    If this.Appended Is Nothing Then Set this.Appended = New BetterArray
End Sub

'@sub-title Register a missing variable and flag reporting.
Public Sub RegisterMissing(ByVal varName As String)
    AddUnique this.Missing, varName
    this.RequiresReport = True
End Sub

'@sub-title Register an updated variable.
Public Sub RegisterUpdated(ByVal varName As String)
    AddUnique this.Updated, varName
End Sub

'@sub-title Register an appended variable and flag reporting.
Public Sub RegisterAppended(ByVal varName As String)
    AddUnique this.Appended, varName
    this.RequiresReport = True
End Sub

'@sub-title Retrieve missing variables as a clone to protect internal state.
Public Property Get MissingVariables() As BetterArray
    Set MissingVariables = CloneArray(this.Missing)
End Property

'@sub-title Retrieve updated variables as a clone.
Public Property Get UpdatedVariables() As BetterArray
    Set UpdatedVariables = CloneArray(this.Updated)
End Property

'@sub-title Retrieve appended variables as a clone.
Public Property Get AppendedVariables() As BetterArray
    Set AppendedVariables = CloneArray(this.Appended)
End Property

'@sub-title Expose whether the merge requires a report.
Public Property Get RequiresReport() As Boolean
    RequiresReport = this.RequiresReport
End Property

'@section Internal Helpers
'===============================================================================

'@sub-title Add a value to the BetterArray only if not already present.
Private Sub AddUnique(ByRef target As BetterArray, ByVal value As String)
    If target Is Nothing Then Set target = New BetterArray
    If LenB(value) = 0 Then Exit Sub
    If Not target.Includes(value) Then
        target.Push value
    End If
End Sub

'@sub-title Return a defensive copy of the supplied BetterArray.
Private Function CloneArray(ByVal source As BetterArray) As BetterArray
    If source Is Nothing Then
        Set CloneArray = New BetterArray
    Else
        Set CloneArray = source.Clone()
    End If
End Function

'@section Interface
'===============================================================================

Private Property Get IDiseaseImportSummary_MissingVariables() As BetterArray
    Set IDiseaseImportSummary_MissingVariables = MissingVariables
End Property

Private Property Get IDiseaseImportSummary_UpdatedVariables() As BetterArray
    Set IDiseaseImportSummary_UpdatedVariables = UpdatedVariables
End Property

Private Property Get IDiseaseImportSummary_AppendedVariables() As BetterArray
    Set IDiseaseImportSummary_AppendedVariables = AppendedVariables
End Property

Private Property Get IDiseaseImportSummary_RequiresReport() As Boolean
    IDiseaseImportSummary_RequiresReport = RequiresReport
End Property

Private Sub IDiseaseImportSummary_RegisterMissing(ByVal varName As String)
    RegisterMissing varName
End Sub

Private Sub IDiseaseImportSummary_RegisterUpdated(ByVal varName As String)
    RegisterUpdated varName
End Sub

Private Sub IDiseaseImportSummary_RegisterAppended(ByVal varName As String)
    RegisterAppended varName
End Sub
