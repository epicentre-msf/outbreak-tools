VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablesSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Analysis")
'@ModuleDescription("analysis tables specifications")
'@IgnoreModule

Option Explicit
Implements ITablesSpecs

Private Type TTablesSpecs
  hRng As Range
  tRng As Range
  lData As ILinelistSpecs
  tType As AnalysisTablesType
End Type


Private this As TTablesSpecs
Private Const CLASSNAME As String = "TableSpecs"

'Instanciation

Public Function Create(headerRng As Range, tableSpecsRng As Range, linelistData As ILinelistSpecs, taType As AnalysisTablesType) As ITablesSpecs
    If headerRng Is Nothing Then ThrowErrorUnableInstanciate "Header range not defined"
    If tableSpecsRng Is Nothing Then ThrowErrorUnableInstanciate "Table range not defined"
    If linelistData Is Nothing Then ThrowErrorUnableInstanciate "Linelist specifications not defined"
    If headerRng.Columns.Count <> tableSpecsRng.Columns.Count Then ThrowErrorUnableInstanciate "Header range and sepicification range don't have same length"
    If headerRng.Row > tableSpecsRng.Row Then ThrowErrorUnableInstanciate "Seems like header comes range after specification range. Maybe you reversed the order? "
    If headerRng.Row < 1 Then ThrowErrorUnableInstanciate "Header range should be at row > 1"

    With New TablesSpecs
      Set .HeaderRange = headerRng
      Set .TableRange = tableSpecsRng
      Set .LinelistSpecifications = linelistData
      .TableType = taType
      Set Create = .Self
    End With
End Function


'Properties
Public Property Get HeaderRange() As Range
  Set HeaderRange = this.hRng
End Property

Public Property Get TableRange() As Range
  Set TableRange = this.tRng
End Property

Public Property Get LinelistSpecifications() As ILinelistSpecs
  Set LinelistSpecifications = this.lData
End Property


Public Property Set HeaderRange(ByVal hRng As Range)
  Set this.hRng = hRng
End Property

Public Property Set TableRange(ByVal tRng As Range)
  Set this.tRng = tRng
End Property

Public Property Set LinelistSpecifications(ByVal lData As ILinelistSpecs)
  Set this.lData = lData
End Property

Public Property Get TableType() As AnalysisTablesType
  TableType = this.tType
End Property

Public Property Let TableType(ByVal taType As AnalysisTablesType)
  this.tType = taType
End Property

Public Property Get Self() As ITablesSpecs
  Set Self = Me
End Property

Private Function ColumnExists(ByVal colName As String) As Boolean
  Dim hRng As Range
  Set hRng = HeaderRange
  ColumnExists = (Not hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False) Is Nothing)
End Function

Private Function ColumnIndex(ByVal colName As String) As Long
  Dim hRng As Range
  Set hRng = HeaderRange
  If ColumnExists(colName) Then
    ColumnIndex = hRng.Find(What:=colName, LookAt:=xlPart, MatchCase:=False).Column - hRng.Column + 1
  Else
    ColumnIndex = -1
  End If
End Function

Private Function Value(ByVal colName As String) As String
  Dim retVal As String
  Dim valIndex As Long
  Dim tRng As Range

  Set tRng = TableRange
  retVal = vbNullString

  If ColumnExists(colName) Then
    valIndex = ColumnIndex(colName)
    retVal = tRng.Cells(1, valIndex).Value
  End If

  Value = retVal
End Function


Private Property Get Data(ByVal rowOrCol As String) As BetterArray
  Dim varsCat As BetterArray
  Dim nameVar As String
  Dim vars As ILLVariables
  Dim lData As ILinelistSpecs

  Set lData = LinelistSpecifications()
  Set vars = LLVariables.Create(lData.Dictionary)
  Set varsCat = New BetterArray
  nameVar = Value(rowOrCol)
  If vars.Contains(nameVar) Then
    Set varsCat = lData.Categories(nameVar)
  End If
  Set Data = varsCat.Clone()
End Property

Private Property Get isNewSection() As Boolean
  Dim sectIndex As Long
  Dim actualSection As String
  Dim previousSection As String
  Dim oldSec As Boolean
  Dim tRng As Range
  Dim hRng As Range
  Dim prevCell As Range

  Set tRng = TableRange()
  Set hRng = HeaderRange()
  oldSec = False

  If ColumnExists("section") Then
    sectIndex = ColumnIndex("section")
    actualSection = Value("section")
    Set prevCell = tRng.Cells(0, sectIndex) '0 for previous line
    previousSection = prevCell.Value
    oldSec = (actualSection = previousSection) And (prevCell.Row <> hRng.Row)
  End If

  isNewSection = (Not oldSec)
End Property

'Errors
'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Interface
Public Property Get ITablesSpecs_TableType() As AnalysisTablesType
  ITablesSpecs_TableType = TableType()
End Property

Public Property Get ITablesSpecs_RowCategories() As BetterArray
  Set ITablesSpecs_RowCategories = Data("row")
End Property

Public Property Get ITablesSpecs_ColumnCategories() As BetterArray
  Set ITablesSpecs_ColumnCategories = Data("column")
End Property

Public Property Get ITablesSpecs_isNewSection() As Boolean
  ITablesSpecs_isNewSection = isNewSection
End Property

Public Function ITablesSpecs_Value(ByVal colName As String) As String
  ITablesSpecs_Value = Value(colName)
End Function
