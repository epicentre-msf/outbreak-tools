VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Analysis Class"


'@Folder("Analysis")
'@ModuleDescription("Analysis Class")
'@IgnoreModule

Option Explicit
Implements ILLAnalysis

Private Type TLLAnalysis
  anash As Worksheet
End Type

Private this As TLLAnalysis
Private Const CLASSNAME As String = "LLAnalysis"
Private Const GLOBALSUMMARY As String = "Tab_global_summary"
Private Const UNIVARIATE As String = "Tab_Univariate_Analysis"
Private Const BIVARIATE As String = "Tab_Bivariate_Analysis"
Private Const TIMESERIES As String = "Tab_TimeSeries_Analysis"
Private Const GRAPHTIMESERIES As String = "Tab_Graph_TimeSeries"
Private Const SPATIAL As String = "Tab_SpactialAnalysis"


'Exposed methods
Public Function Create(ByVal anaWksh As Worksheet) As ILLAnalysis

   If anaWksh Is Nothing Then ThrowErrorUnableInstanciate "Worksheet object not defined"
   With anaWksh.Cells(1, 1)
    If .Font.Color <> vbBlue Or .Value <> "analysis" Then ThrowErrorUnableInstanciate "Not a valid analysis worksheet"
   End With

  With New LLAnalysis
    Set .Wksh = anaWksh
    Set Create = .Self
  End With
End Function


Public Property Get Wksh() As Worksheet
 Set Wksh = this.anash
End Property

Public Property Set Wksh(ByVal anash As Worksheet)
  Set this.anash = anash
End Property

Public Property Get Self() As ILLAnalysis
    Set Self = Me
End Property


'Find if a column Exists in a listObject of the analysis
Private Function ColumnExists(ByVal Lo As ListObject, ByVal colName As String) As Boolean
  Dim hRng As Range
  Set hRng = Lo.HeaderRowRange

  ColumnExists = (Not hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal Lo As ListObject, ByVal colName) As Long
  Dim Index As Long
  Dim hRng As Range

  Set hRng = Lo.HeaderRowRange
  If ColumnExists(Lo, colName) Then
    Index = hRng.Find(What:=colName, LookAt:=xlWhole, MatchCase:=False).Column - hRng.Column + 1
  Else
    Index = -1
  End If
  ColumnIndex = Index
End Function

Private Function DataRange(ByVal Lo As ListObject, ByVal colName As String) As Range
  If ColumnExists(Lo, colName) Then Set DataRange = Lo.ListColumns(ColumnIndex(Lo, colName)).Range
End Function


Private Sub Import(ByVal impWksh As Worksheet)

  Dim anash As Worksheet
  Dim impRng As Range
  Dim actRng As Range
  Dim impAddress As String
  Dim impLo As ListObject

  Set anash = Wksh()

  'Test if the sheet is empty
  If impWksh Is Nothing Then Exit Sub 'Add trace for import

  anash.Cells.Clear

  On Error GoTo LoFail
    Set impLo = impWksh.ListObjects(1)
  On Error GoTo 0


  For Each impLo In impWksh.ListObjects

    Set impRng = impLo.Range
    impAddress = impRng.Address
    Set actRng = anash.Range(impAddress)
    actRng.Value = impRng.Value

    actRng.Cells(-1, 1).Value = impRng.Cells(-1, 1).Value
    actRng.Cells(-1, 1).Font.Size = impRng.Cells(-1, 1).Font.Size
    actRng.Cells(-1, 1).Font.Color = impRng.Cells(-1, 1).Font.Color
    actRng.Cells(-1, 1).Font.Bold = True

    'Define the listobject
    With anash
      .ListObjects.Add(xlSrcRange, actRng, , xlYes).Name = impLo.Name
      .ListObjects(impLo.Name).TableStyle = impLo.TableStyle
    End With
  Next

  anash.Cells(1, 1).Value = "analysis"
  anash.Cells(1, 1).Font.Color = vbBlue

  Exit Sub
LoFail:
  'add trace for no list objects
  Exit Sub
End Sub

'Translate
Private Sub Translate(ByVal transObject As ITranslation)
  Dim transCols As BetterArray
  Dim anash As Worksheet
  Dim counter As Long
  Dim Lo As ListObject
  Dim colRng As Range

  Set transCols = New BetterArray
  Set anash = Wksh()
  transCols.Push "section", "title", "summary label"

  For Each Lo In anash.ListObjects
    For counter = transCols.LowerBound To transCols.UpperBound
      If ColumnExists(Lo, transCols.Item(counter)) Then
        Set colRng = DataRange(Lo, transCols.Item(counter))
        transObject.TranslateRange colRng
      End If
    Next
    'Formulas
    If ColumnExists(Lo, "summary function") Then
      Set colRng = DataRange(Lo, "summary function")
      transObject.TranslateRange colRng, containsFormula:=True
    End If
  Next
End Sub



'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub


'Interface
Public Sub ILLAnalysis_Import(ByVal impWksh As Worksheet)
  Import impWksh
End Sub

Public Sub ILLAnalysis_Translate(ByVal transObject As ITranslation)
    Translate transObject
End Sub

Public Property Get ILLAnalysis_Wksh() As Worksheet
    Set ILLAnalysis_Wksh = Wksh()
End Property
