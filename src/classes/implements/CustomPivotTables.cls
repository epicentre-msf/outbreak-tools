VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "CustomPivotTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Custom Pivot Tables for the linelist"

'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Custom Pivot Tables for the linelist")

Option Explicit
Implements ICustomPivotTables

Private Const CLASSNAME As String = "CustomPivotTables"
Private Const PIVOTPREFIX As String = "PivotTable_"

Private Type TCustPivot
  sh As Worksheet
End Type

Private this As TCustPivot
Private PIVOTCOUNTER As Long

Public Function Create(ByVal sh As Worksheet) As ICustomPivotTables

  If sh Is Nothing Then ThrowError "Unable to create custom pivot, sheet not found"
  
  With New CustomPivotTables
    Set .Wksh = sh 
    Set Create = .Self
  End With
End Function

Public Property Get Self() As ICustomPivotTables
  Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property

'Initialize where to write output range
Private Sub InitializeOutputAddress()
  Dim sh As Worksheet
  Set sh = Wksh()
  If IsEmpty(sh.Cells(1, 1)) Then
    sh.Cells(1, 1).Value = "$B$2"
    sh.Cells(1, 1).Font.Color = vbWhite
    pivotCounter = 1
  End If
End Sub

'Get the output Range of the pivot table
Private Property Get OutputRange() As Range
  Dim sh As Worksheet

  Set sh = Wksh()
  InitializeOutputAddress
  Set OutputRange = sh.Range(sh.Cells(1, 1).Value).Offset(6)
End Property

'Add a title for the pivot table
Private Sub AddTitle(ByVal title As String, Byval pivotName As String)
  Dim sh As Worksheet
  Dim titleRng As Range

  Set sh = Wksh()
  
  InitializeOutputAddress
  
  Set titleRng = sh.Range(sh.Cells(1, 1).Value).Offset(2)
  titleRng.Value = pivotName & " " & PIVOTCOUNTER & " - " & title
  titleRng.Font.Color = RGB(155, 109, 88)
  titleRng.Font.Italic = True
End Sub

'Add a new pivot table
'-pivotName: Name of a pivot table depending on the language
'-title: Name of the title of the pivot table
'-tablename: name of the pivot table (instead of using pivot table 1, 2, etc)
Private Sub AddTable(ByVal title As String, ByVal tablename As String, _ 
                     Optional ByVal pivotName As String = "Pivot Table")

  Dim outrng As Range
  Dim sh As Worksheet
  Dim rng As Range
  
  Set sh = Wksh()
  Set outrng = OutputRange()

  On Error GoTo ErrAdding
    'Add the pivot table before adding the title and other elements
    'If it fails, the title is not added
    sh.PivotTableWizard SourceType:=xlDatabase, SourceData:=tablename, _
                      TableName:=PIVOTPREFIX & tablename, _ 
                      TableDestination:=outrng
  On Error GoTo 0

  AddTitle title, pivotName
  'Change the new output address to another one
  Set rng = sh.Range(sh.Cells(1, 1).Value)
  sh.Cells(1, 1).Value = outrng.Offset(60).Address
  PIVOTCOUNTER = PIVOTCOUNTER + 1
  Exit Sub
ErrAdding:
  Debug.Print "Unable to add pivot table " & tablename & ":" & Err.Number & "-" & Err.Description  
End Sub


'Errors management
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Interface

Private Sub ICustomPivotTables_Add(ByVal title As String, _ 
                                   ByVal tablename As String, _ 
                                   Optional ByVal pivotName As String = "Pivot Table")
  AddTable title, tableName, pivotName
End Sub

'Exposed methods