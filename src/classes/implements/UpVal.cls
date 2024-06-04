VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Watch updated values in the linelist"


'@IgnoreModule ReadOnlyPropertyAssignment, UseMeaningfulName, HungarianNotation
'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Watch updated values in the linelist")

Option Explicit

Implements IUpVal

'Exposed methods
Private Type TUpdatedVal
  sh As Worksheet
End Type

Private this As TUpdatedVal

Private Const CLASSNAME As String = "UpVal"
Private Const RNGLISTAUTOUPDATED As String = "RNG_UpdateListAuto"
Private Const RNGIMPORTDONE As String = "RNG_ImportDone"
Private Const RNGGEOBASEMANUALUPDATE As String = "RNG_GeoUpdateManual"
Private Const RNGSHOWALLOPTIONALS As String = "RNG_ShowAllOptionals"
Private Const RNGEXPORTINIT As String = "RNG_ExportInit"
Private Const RNGNUMBEXP As String = "RNG_NumberExports"

Private UPDATERANGESLIST As BetterArray

Public Function Create(ByVal sh As Worksheet) As IUpVal

  CheckRequirements sh
  
  With New UpVal
    Set .Wksh = sh
    Set Create = .Self()
  End With

End Function

Public Property Get Self() As IUpVal
  Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

Private Sub InitalizeRanges()
  If UPDATERANGESLIST Is Nothing Then
    Set UPDATERANGESLIST = New BetterArray
    UPDATERANGESLIST.Push RNGLISTAUTOUPDATED, RNGIMPORTDONE, RNGGEOBASEMANUALUPDATE, _ 
                          RNGSHOWALLOPTIONALS
  End If
End Sub

'Retrieve value from an updated range

Private Property Get Value(ByVal rngName As String) As String
  Dim retVal As String
  Dim sh As Worksheet

  Set sh = Wksh()
  InitalizeRanges

  On Error Resume Next
  If  UPDATERANGESLIST.Includes(rngName) Then retVal = CStr(sh.Range(rngName).Value)
  On Error GoTo 0

  Value = retVal
End Property

Private Sub SetValue(ByVal rngName As String, newValue As String)
  Dim sh As Worksheet
  
  Set sh = Wksh()
  InitalizeRanges
  
  If UPDATERANGESLIST.Includes(rngName) Then sh.Range(rngName).Value = newValue

End Sub

'Checking for errors at instanciation
'Errors checking at instanciation

Private Sub CheckRequirements(ByVal sh As Worksheet)
  
  Dim counter As Long
  Dim rngName As String
  Dim rngArray As BetterArray
  Set rngArray = New BetterArray
  If sh Is Nothing Then ThrowErrorUnableInstanciate "Update values worksheet not defined"

  rngArray.Push RNGLISTAUTOUPDATED, RNGIMPORTDONE, RNGGEOBASEMANUALUPDATE, RNGSHOWALLOPTIONALS, _ 
                RNGEXPORTINIT, RNGNUMBEXP

  For counter = rngArray.LowerBound To rngArray.UpperBound
    rngName = rngArray.Item(counter)
    If Not RangeExists(sh, rngName) Then _
     ThrowErrorUnableInstanciate " Range " & rngName & " Not found!"
  Next
  
End Sub

'Errors managements
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, _
   "Unable to create an instance of the class " & CLASSNAME & ": " & errorMessage
End Sub

Private Function RangeExists(ByVal sh As Worksheet, ByVal rngName As String) As Boolean
  Dim rng As Range

  On Error Resume Next
    Set rng = sh.Range(rngName)
  On Error GoTo 0

  RangeExists = Not (rng Is Nothing)
End Function

'Set and Get updated values
Private Property Get IUpVal_Value(ByVal rngName As String) As String
  IUpVal_Value = Value(rngName)
End Property

Private Sub IUpVal_SetValue(ByVal rngName As String, ByVal newValue As String)
  SetValue rngName, newValue
End Sub
