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


'@IgnoreModule ReadOnlyPropertyAssignment : Ignore when you affect property Value of a range object
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

'Retrieve value from an updated range

Private Property Get Value(ByVal rngName As String) As String
  Dim retVal As String
  Dim sh As Worksheet

  Set sh = Wksh()

  If (rngName = RNGLISTAUTOUPDATED) Or _
     (rngName = RNGIMPORTDONE) Or _
     (rngName = RNGGEOBASEMANUALUPDATE) Then
     On Error Resume Next
      retVal = CStr(sh.Range(rngName).Value)
     On Error GoTo 0
  End If

  Value = retVal

End Property

Private Sub SetValue(ByVal rngName As String, newValue As String)

  Dim sh As Worksheet
  Set sh = Wksh()
  
  If (rngName = RNGLISTAUTOUPDATED) Or _
     (rngName = RNGIMPORTDONE) Or _
     (rngName = RNGGEOBASEMANUALUPDATE) Then
    sh.Range(rngName).Value = newValue
  End If

End Sub

'Checking for errors at instanciation
'Errors checking at instanciation

Private Sub CheckRequirements(ByVal sh As Worksheet)
  
  Dim counter As Long
  Dim rngName As String
  Dim rngArray As BetterArray
  Set rngArray = New BetterArray
  If sh Is Nothing Then ThrowErrorUnableInstanciate "Update values worksheet not defined"

  rngArray.Push RNGLISTAUTOUPDATED, RNGIMPORTDONE, RNGGEOBASEMANUALUPDATE

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
