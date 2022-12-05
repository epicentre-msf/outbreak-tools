VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LLSpatial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'@Folder("Linelist")
'@ModuleDescription("Manage the Spatial analysis tables")
'@IgnoreModule
Implements ILLSpatial

Private Type TLLSpatial
    sh As Worksheet
End Type

Private this As TLLSpatial
Private Const CLASSNAME As String = "LLSpatial"
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const ADMIN As String = "adm"
Private Const CONCAT As String = "concat_"

Public Function Create(ByVal sh As Worksheet) As ILLSpatial
    
    'output worksheet to add the spatial tables
    CheckRequirements sh

    With New LLSpatial
        Set .Wksh = sh
        Set Create = .Self
    End With

End Function


Public Property Get Self() As ILLSpatial
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.sh
End Property

Public Property Set Wksh(ByVal sh As Worksheet)
  Set this.sh = sh
End Property


Private Property Get CurrentHeaderName(ByVal varName As String, ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

Private Property Get CurrentLoName(ByVal varName As String, ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'Test if a listobject exists in a sheet
Private Function Exists(ByVal varName As String) As Boolean
  Dim counter As Long
  Dim testExistence As Boolean

  For counter = 1 To 4
    loName = CurrentLoName(varName, ADMIN & counter)
    testExistence = LoExists(loName)
    If Not testExistence Then Exit For
  Next

  Exits = testExistence
End Function

Private Function LoExists(ByVal loName As String) As Boolean
  Dim sh As Worksheet
  Dim Lo As ListObject
  Set sh = Wksh()

  On Error Resume Next
    Set Lo = sh.ListObjects(loName)
  On Error GoTo 0

  LoExists = (Not Lo Is Nothing)
End Function

'Add conditions for the formulas


Private Sub Update()
  Dim actSh As Worksheet 'Actual worksheet (the spatial worksheet)
  Dim sh As Worksheet 'All worksheets in the workbook
  Dim filtsh As Worksheet
  Dim wb As Workbook
  Dim rng As Range
  Dim tableName As String
  Dim varName As String
  Dim cellRng As Range
  Dim controlValue As String
  Dim Lo As ListObject

  Set actSh = Wksh()
  Set wb = actSh.Parent

  For Each sh In wb.Worksheets

    'Test if it is a HList
    If sh.Cells(1, 3) = "HList" Then

      Set filtsh = wb.Worksheets(sh.Cells(1, 5).Value)

      'Exit the sub if there is no data in the filters
      If Not (filtsh.ListObjects(1).DataBodyRange Is Nothing) Then
          
         tableName = sh.Cells(1, 4).Value
          Set cellRng = sh.Range(tableName & "_" & "START")
          Set Lo = filtsh.ListObjects(1)

          'Offset - 1 is the range with the variable names
          Do While Not IsEmpty(cellRng.Offset(-1))
            controlValue = cellRng.Offset(-5).Value

            If controlValue = "geo1" Then

              varName = Replace(ADMIN & "1_", cellRng.Offset(-1).Value, "")

              If Exists(varName) Then
                Set rng = Lo.ListColumns(CONCAT & ADMIN & "1_" & varName).DataBodyRange()
                
                UpdateSpatialValue rng, varName

              End If

            End If
          Loop
      End If
    End If
  Next
End Sub


Private Sub UpdateSpatialValue(ByVal inputRng As Range, ByVal varName As String)

  Dim pasteRng As Range
  Dim counter As Long
  Dim sh As Worksheet
  Dim cellRng As Range
  Dim rng As Range
  Dim Lo As ListObject
  Dim loName As String
  Dim adminName As String
  Dim sortRng As Range
  Dim nbRows As Long

  Set sh = Wksh()
  Set rng = inputRng
  nbRows = rng.Cells.Count

  For counter = 1 To 4
      loName = CurrentLoName(varName, ADMIN & counter)
      Set Lo = sh.ListObjects(loName)

      'First cell where to paste the values of the range
      Set cellRng = Lo.Range.Cells(2, 1)
      Set pasteRng = sh.Range(cellRng.Cells(1, 1), cellRng.Cells(nbRows, 1))

      'copy values to the pasting Rng
      pasteRng.Value = rng.Value

      'Sort on the formula column
      Set sortRng = Lo.ListColumns("formula_" & ADMIN & "_" & counter).DataBodyRange
      Lo.DataBodyRange.Sort key1:=sortRng, order1:=xlDescending
      Set rng = rng.Offset(, 1)
  Next
End Sub

Private Property Get FindTopValue(ByVal adminName As String, ByVal orderValue As Integer, ByVal varName As String) As String
  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim loName As String
  Dim rng As Range
  Set sh = Wksh()

  loName = CurrentLoName(varName, adminName)
  If Not LoExists(loName) Then Exit Property
  Set Lo = sh.ListObjects(loName)
  Set rng = Lo.ListColumns(1).DataBodyRange()
  FindTopValue = rng.Cells(orderValue, 1).Value

End Property


Private Sub CheckRequirements(ByVal sh As Worksheet)

  If sh Is Nothing Then ThrowError "Worksheet not found, please provide a worksheet"
  If sh.Name <> SPATIALSHEET Then ThrowError "The worksheet name is not " & SPATIALSHEET

End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Public Property Get ILLSpatial_Exists(ByVal varName As String) As Boolean
    ILLSpatial_Exists = Exists(varName)
End Property

Public Sub ILLSpatial_Update()
  Update
End Sub


Public Property Get ILLSpatial_FindTopValue(ByVal adminName As String, ByVal levelValue As Integer, ByVal varName As String) As String
  ILLSpatial_FindTopValue = FindTopValue(adminName, levelValue, varName)
End Property
