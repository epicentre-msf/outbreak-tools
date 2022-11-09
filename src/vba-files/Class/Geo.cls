VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Geo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Geobase management class")
'@IgnoreModule

Option Explicit
Implements IGeo

Private Type TGeo
  geosh As Worksheet
End Type

'Exposed methods
Private this As TGeo

Public Function Create(ByVal sh As Worksheet, admTableNames As BetterArray, geoRngName As String) As IGeo

    CheckRequirements sh, admTableNames

End Function

Private Property Get Wksh() As Worksheet
  Set Wksh = this.geosh
End Property

Private Property Set Wksh(ByVal sh As Worksheet)
  Set this.geosh = sh
End Property

Private Sub CheckRequirements(ByVal sh As Worksheet, ByVal admTableNames As BetterArray)


End Sub





Private Sub RangeExists(ByVal rngName As String)
    Dim sh As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set sh = Wksh()
    Set rng = sh.Range(rngName)
    If rng Is Nothing Then
        ThrowError "Range " & rngName & " does not exists in worksheet " & sh.Name
    End If
    On Error GoTo 0
End Sub

Private Sub LoExists(ByVal loNames As BetterArray)
    Dim Lo As ListObject
    Dim sh As Worksheet
    Dim counter As Long
    Dim lName As String


    For counter = loNames.LowerBound To loNames.UpperBound
        lName = loNames.Item(counter)
        On Error Resume Next
        Set Lo = sh.ListObjects(loNames)
        If Lo Is Nothing Then
            ThrowError "Table " & lName & " does not exists in worksheet" & sh.Name
        End If
        On Error GoTo 0
    Next

End Sub


Public Sub Export(Wkb As Workbook)

End Sub


Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub
