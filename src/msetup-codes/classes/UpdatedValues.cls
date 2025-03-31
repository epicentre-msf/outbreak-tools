VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdatedValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Watch updated values in the setup"


'@PredeclaredId
'@Folder("setup file")
'@ModuleDescription("Watch updated values in the setup")
Option Explicit
Implements IUpdatedValues

Private Const CLASSNAME As String = "IUpdatedValues"
Private Const TRANSLATETEXT As String = "translate as text"
Private Const TRANSLATEFORM As String = "translate as formula"
Private Const WATCHUPDATE As String = "watch for update"

Private Type TUpValues
  upLo As ListObject
  sh As Worksheet
  upName As String
End Type

Private this As TUpValues

Public Function Create(ByVal sh As Worksheet, ByVal upName As String) As IUpdatedValues
  If sh Is Nothing Then ThrowError "Unable to create an instance of the class UpdatedValues, the worksheet does not exists"
  If upName = vbNullString Then ThrowError "Please provide a name for your object for UpdatedValues"

  With New UpdatedValues
    Set .Wksh = sh
    .ID = upName
    Set Create = .Self()
  End With
End Function

Public Property Get Self() As IUpdatedValues
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
    Set this.sh = sh
End Property

Public Property Get ID() As String
  ID = this.upName
End Property

Public Property Let ID(ByVal upName As String)
  this.upName = upName
End Property

Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()
    colIndex = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column + 2
    LastColumn = colIndex
End Property

Private Property Get Table() As ListObject
  Dim sh As Worksheet
  Dim colIndex As Long
  Dim rng As Range
  Set sh = Wksh()

  If this.upLo Is Nothing Then
    If Not Exists(ID) Then
      colIndex = LastColumn()
      With sh
        .Cells(1, colIndex).Value = "colname"
        .Cells(1, colIndex + 1).Value = "rngname"
        .Cells(1, colIndex + 2).Value = "updated"
        .Cells(1, colIndex + 3).Value = "headername"
        Set rng = .Range(.Cells(1, colIndex), .Cells(2, colIndex + 3))
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = CurrentLoName(ID)
      End With
    End If
    Set this.upLo = sh.ListObjects(CurrentLoName(ID))
  End If
  Set Table = this.upLo
End Property

Private Property Get IsUpdated(ByVal checkColName As String) As Boolean
  Dim upLo As ListObject
  Dim rng As Range
  Dim cellRng As Range
  Dim rngName As String

  Set upLo = Table()

  rngName = "RNG_" & LCase(Application.WorksheetFunction.Trim(Replace(checkColName, " ", "_"))) & "_" & ID()

  Set rng = upLo.ListColumns("rngname").Range
  Set cellRng = rng.Find(What:=rngName, lookAt:=xlWhole, MatchCase:=True)
  If cellRng Is Nothing Then Exit Property
  IsUpdated = (cellRng.Cells(1, 2).Value = "yes")
End Property

Private Sub AddColumns(ByVal Lo As ListObject)

  Dim hRng As Range 'header row range of the listObject
  Dim idRng As Range
  Dim upLo As ListObject 'Current Updating listObject
  Dim wb As Workbook
  Dim sh As Worksheet
  Dim rngName As String
  Dim colName As String
  Dim headerName As String
  Dim idName As String
  Dim counter As Long

  If Lo Is Nothing Then ThrowError "ListObject is not initialize, please provide valid listObject"
  Set upLo = Table()

  'Check if the column Exists in the header range
  Set hRng = Lo.HeaderRowRange

  For counter = 1 To hRng.Columns.Count

    colName = hRng.Cells(1, counter).Value
    headerName = hRng.Cells(-2, counter).Value

    If (headerName = TRANSLATETEXT) Or _
       (headerName = TRANSLATEFORM) Or _
       (headerName = WATCHUPDATE) Then
      idName = LCase(Application.WorksheetFunction.Trim(Replace(colName, " ", "_")))
      'Check if the id already Exists, in that case exit
      Set idRng = upLo.ListColumns("rngname").Range()

      If (idRng.Find(What:=idName, lookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
          Set sh = Wksh()
          Set wb = sh.Parent

          'Now add values and names to the workbook
          rngName = "RNG_" & idName & "_" & ID()
          On Error Resume Next
              wb.Names(rngName).Delete
          On Error GoTo 0

          wb.Names.Add Name:=rngName, RefersToR1C1:="=" & Lo.Name & "[" & colName & "]"
          'Write values to the listObject of the update
          WriteValues Lo.Name & "-" & colName, rngName, headerName
      End If
    End If
  Next
End Sub

Private Sub WriteValues(ByVal colName As String, ByVal rngName As String, _
                        ByVal headerName As String, _
                        Optional ByVal statusType As String = "no")
  Dim upLo As ListObject
  Dim rowRng As Range
  Dim loRng As Range
  Dim sh As Worksheet

  Set upLo = Table()
  Set rowRng = upLo.HeaderRowRange.Cells(1, 1)

  Do While (rowRng.Value <> vbNullString)
    Set rowRng = rowRng.Offset(1)
  Loop

  rowRng.Cells(1, 1).Value = colName
  rowRng.Cells(1, 2).Value = rngName
  rowRng.Cells(1, 3).Value = statusType
  rowRng.Cells(1, 4).Value = headerName

  'Resize the listObject
  Set sh = Wksh()
  Set loRng = sh.Range(upLo.HeaderRowRange.Cells(1, 1), rowRng.Cells(1, 4))
  upLo.Resize loRng
End Sub

Private Sub WriteUpdate(ByVal pos As Long, Optional ByVal statusType As String = "yes")
  Dim upLo As ListObject
  Dim idRng As Range
  'Find the colName in the listObject
  Set upLo = Table()
  Set idRng = upLo.ListColumns("rngname").Range
  idRng.Cells(pos, 2).Value = statusType
End Sub

Private Sub ClearUp()
  Dim upLo As ListObject
  Set upLo = Table()
  If Not upLo.DataBodyRange Is Nothing Then upLo.DataBodyRange.Delete
End Sub

Private Sub DeleteUp()
  Dim upLo As ListObject
  Dim rng As Range
  Set upLo = Table()
  Set rng = upLo.Range
  If Not upLo.DataBodyRange Is Nothing Then upLo.DataBodyRange.Delete
  upLo.Unlist
  rng.Cells.Clear
End Sub

'Check the update status
Private Sub CheckStatus(ByVal sh As Worksheet, ByVal Target As Range)
  Dim rng As Range
  Dim counter As Long
  Dim idRng As Range
  Dim upLo As ListObject
  Dim rngName As String

  If Not Exists(ID) Then ThrowError "Unable to add elements to a non existant update table, check the id name"
  Set upLo = Table()
  Set idRng = upLo.ListColumns("rngname").Range
  For counter = 1 To idRng.Rows.Count
    rngName = idRng.Cells(counter, 1).Value
    On Error Resume Next
      Set rng = sh.Range(rngName)
    On Error GoTo 0
    If Not (rng Is Nothing) Then
      If Not IsUpdated(rngName) Then
        If Not (Intersect(Target, rng) Is Nothing) Then
          WriteUpdate counter
          Exit Sub
        End If
      End If
      Set rng = Nothing
    End If
  Next
End Sub

Private Sub ResetUpdateStatus()
  Dim upLo As ListObject
  Dim idRng As Range
  Dim counter As Long
  Set upLo = Table()
  If upLo.DataBodyRange Is Nothing Then Exit Sub
  Set idRng = upLo.ListColumns("rngname").DataBodyRange
  For counter = 1 To idRng.Rows.Count
    idRng.Cells(counter, 2).Value = "no"
  Next
End Sub

Private Property Get CurrentLoName(ByVal listName As String) As String
    CurrentLoName = "UpLo_" & listName
End Property

'Test if a listobject exists in a sheet
Public Function Exists(ByVal listName As String) As Boolean
    Dim Lo As ListObject
    Dim loname As String
    Dim sh As Worksheet
    Set sh = Wksh()
    loname = CurrentLoName(listName)
    On Error Resume Next
    Set Lo = sh.ListObjects(loname)
    On Error GoTo 0
    Exists = (Not Lo Is Nothing)
End Function

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Private Sub IUpdatedValues_AddColumns(ByVal Lo As ListObject)
  AddColumns Lo
End Sub

Private Property Get IUpdatedValues_IsUpdated(ByVal colName As String) As Boolean
  IUpdatedValues_IsUpdated = IsUpdated(colName)
End Property

Private Sub IUpdatedValues_CheckUpdate(ByVal sh As Worksheet, ByVal Target As Range)
  CheckStatus sh, Target
End Sub

Private Sub IUpdatedValues_ClearUp()
  ClearUp
End Sub

Private Sub IUpdatedValues_DeleteUp()
  DeleteUp
End Sub
