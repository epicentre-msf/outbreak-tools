VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinelistSheetCreationStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("Tests.Stubs")
'@ModuleDescription("Preparation step stub that creates configured sheets on the linelist")
'@IgnoreModule UnrecognizedAnnotation, UseMeaningfulName, ProcedureNotUsed

Option Explicit

Implements ILinelistPreparationStep

Private Type TStep
    sheetNames As BetterArray
End Type

Private this As TStep

Public Sub Initialise(ParamArray names() As Variant)
    Dim index As Long
    Dim lower As Long
    Dim upper As Long

    Set this.sheetNames = New BetterArray
    this.sheetNames.LowerBound = 1

    On Error GoTo Empty
        lower = LBound(names)
        upper = UBound(names)
        For index = lower To upper
            this.sheetNames.Push CStr(names(index))
        Next index
        Exit Sub
Empty:
    Err.Clear
End Sub

Private Sub ILinelistPreparationStep_Execute(ByVal context As ILinelistPreparationContext)
    Dim index As Long

    If this.sheetNames Is Nothing Then Exit Sub

    For index = this.sheetNames.LowerBound To this.sheetNames.UpperBound
        context.Linelist.AddOutputSheet CStr(this.sheetNames.Item(index))
    Next index
End Sub

