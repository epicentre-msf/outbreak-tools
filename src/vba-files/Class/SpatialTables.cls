VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpatialTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'@Folder("Linelist")
'@ModuleDescription("Manage the Spatial analysis tables")
'@IgnoreModule
Option Explicit

Implements ISpatialTables

Private Type TSpatialTables
    sh As Worksheet
    tab As ICrossTable
    condVars As BetterArray
    condTests As BetterArray
End Type

Private this As TSpatialTables
Private Const CLASSNAME As String = "SpatialTables"
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const ADMIN As String = "adm"
Private Const CONCAT As String = "concat_"

Public Function Create(ByVal tabl As ICrossTable) As ISpatialTables
    
    'output worksheet to add the spatial tables
    CheckRequirements tabl

    With New SpatialTables
        Set .Table = tabl
        Set Create = .Self
    End With

End Function

Public Property Get Self() As ISpatialTables
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Dim tabl As ICrossTable
  Dim wb As Workbook
  Dim sh As Worksheet
  
  If this.sh Is Nothing Then
    Set tabl = Table()
    Set sh = tabl.Wksh()
    Set wb = sh.Parent
    Set this.sh = wb.Worksheets(SPATIALSHEET)
  End If

  Set Wksh = this.sh

End Property

Public Property Get Table() As ICrossTable
  Set Table = this.tab
End Property

Public Property Set Table(ByVal tabl As ICrossTable)
  Set this.tab = tabl
End Property

Private Property Get LastColumn() As Long
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()
    With sh
        colIndex = .Cells(1, .Columns.Count).End(xlToLeft).Column
        colIndex = IIf(colIndex = 1, 7, colIndex + 2)
    End With
    LastColumn = colIndex
End Property

Private Property Get CurrentListName(ByVal varName As String) As String
    CurrentListName = Replace(Application.WorksheetFunction.Trim(varName), " ", "_")
End Property

Private Property Get CurrentHeaderName(ByVal varName As String, ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & CurrentListName(varName)
End Property

Private Property Get CurrentLoName(ByVal varName As String, ByVal adminName As String) As String
    CurrentLoName = "spatial_" & adminName & "_" & CurrentListName(varName)
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

'All the formula management from crossTable
Private Function FailedFormula(ByVal frmName As String) As Boolean

  Dim prevAlert As Long
  Dim sh As Worksheet
  Dim formTest As String
  
  prevAlert = Application.DisplayAlerts

  Set sh = Wksh
  With sh
      Application.DisplayAlerts = False
      .Cells(1, 1).Clear
      On Error GoTo Fail
      
      'Test the formula on one cell
      .Cells(1, 1).FormulaArray = "= " & frmName
      
      .Cells(1, 1).Clear
      FailedFormula = False
      Exit Function
  End With
  
  Application.DisplayAlerts = prevAlert
Fail:
  FailedFormula = True
  Application.DisplayAlerts = prevAlert
End Function

Private Function EmpStr() As String
  EmpStr = Chr(34) & Chr(34)
End Function

Private Function NonEmpStr() As String
  NonEmpStr = Chr(34) & "<>" & Chr(34)
End Function

'Add conditions for the formulas
Private Sub AddTableConditions(varName As String, conditionName As String)

  If this.condVars Is Nothing Then
      Set this.condVars = New BetterArray
      this.condVars.LowerBound = 1
  End If

  If this.condTests Is Nothing Then
      Set this.condTests = New BetterArray
      this.condTests.LowerBound = 1
  End If

  If varName <> vbNullString Then
      this.condVars.Push varName
      this.condTests.Push conditionName
  End If
  
End Sub

Private Sub ClearTableConditions(Optional ByVal removeNumber As Integer = 0)
  Dim counter As Integer

  If this.condVars Is Nothing Then
      Set this.condVars = New BetterArray
      this.condVars.LowerBound = 1
      removeNumber = 0
  End If

  If this.condTests Is Nothing Then
      Set this.condTests = New BetterArray
      this.condTests.LowerBound = 1
      removeNumber = 0
  End If

  If removeNumber <= 0 Then
      this.condVars.Clear
      this.condTests.Clear
  Else
      For counter = 1 To removeNumber
          this.condVars.Pop
          this.condTests.Pop
      Next
  End If
End Sub

Private Property Get CurrentExcelFormula(tabFormObject As IFormulas, _
                                        Optional ByVal scope As ExcelFormulaScope = formulaScopeanalysis, _
                                         Optional ByVal prefix As String = "f") As String
  'use the formula
  Dim cond As IFormulaCondition
  Dim eform As String

  If Not tabFormObject.Valid() Then Exit Property
     
  If scope = 1 Then
      eform = tabFormObject.ParsedLinelistFormula(useTableName:=True, tablePrefix:=prefix)
  Else
      If this.condVars Is Nothing Or this.condTests Is Nothing Then Exit Property
      Set cond = FormulaCondition.Create(this.condVars, this.condTests)
      eform = tabFormObject.ParsedAnalysisFormula(cond, tablePrefix:=prefix)
  End If

  'Test here the formula before assigining i
  If FailedFormula(eform) Then eform = vbNullString

  CurrentExcelFormula = eform
End Property

'Create require listObjects to interact with the worksheets
Private Sub Prepare()

  Dim rng As Range
  Dim sh As Worksheet

  Set sh = Wksh()

  With sh
    Set rng = .Range(.Cells(1, 3), .Cells(2, 3))
    rng.Cells(1, 1).Value = "listofvars"
    .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "listofgeovars"

    Set rng = .Cells(1, 5)
    rng.Name = "RNG_PastingCol" 'Pasting column

    Set rng = .Cells(1, 1)
    rng.Name = "RNG_TestingFormula" 'Testing formula range (just for readbility purposes)
  End With
End Sub

'Add New list to the dropdown sheet
Private Sub Add(ByVal formObject As IFormulas)
  
    Dim sh As Worksheet
    Dim colIndex As Long
    Dim varName As String
    Dim headerName As String                     'header Name
    Dim loName As String                         'List object name
    Dim rng As Range
    Dim wkb As Workbook
    Dim counter As Long
    Dim specs As ITablesSpecs
    Dim tabl As ICrossTable
    Dim cellRng As Range
    Dim rng As Range
    Dim finalFormula As String
    Dim adminName As String

    Set tabl = Table
    Set specs = tabl.Specifications
    varName = specs.Value("row")

    If Exists(varName) Then Exit Sub
    If Not LoExists("listofgeovars") Then Prepare

    Set sh = Wksh()
      
    With sh

      For counter = 1 To 4
        
        ClearTableConditions

        loName = CurrentLoName(varName, ADMIN & counter)
        colIndex = LastColumn()
        
        'Add the table names for admin 1, only admin1, for admin2, admin1 and admin2, etc...
        adminName = ADMIN & counter
        headerName = CurrentHeaderName(varName, adminName)
        Set cellRng = .Cells(1, colIndex)
        cellRng.Value = headerName
        Set cellRng = cellRng.Offset(1)
        AddTableConditions CONCAT & adminName & "_" & varName, "= " & cellRng.Address(Rowabsolute:=False)

        'Formula value
        Set cellRng = cellRng.Offset(0, 1) 'formula header
        cellRng.Value = "formula_" & ADMIN & "_" & counter
        
        'Where to write the formula
        Set cellRng = cellRng.Offset(1)
        
        'final formula
        finalFormula = CurrentExcelFormula(formObject)

        'Test final formula to see if it returns something
        If finalFormula <> vbNullString Then cellRng.FormulaArray = "= " & finalFormula

        'Define the listObject
        Set rng = .Range(.Cells(1, colIndex), .Cells(2, colIndex + 1))

        'Add A list object table
        .ListObjects.Add(xlSrcRange, rng, , xlYes).Name = loName
      Next

    End With
End Sub

Private Sub CheckRequirements(ByVal tabl As ICrossTable)
  
  Dim wb As Workbook
  Dim sh As Worksheet
  
  If tabl Is Nothing Then ThrowError "No Table found, please provide a table for instanciation"
  Set sh = tabl.Wksh
  Set wb = sh.Parent

  sheetExists wb, SPATIALSHEET
End Sub

Private Sub sheetExists(ByVal specswkb As Workbook, ByVal sheetName As String)
    Dim sh As Worksheet                          'Just try to set the workbook if it fails it is closed
    Dim checkExists As Boolean
    
    On Error Resume Next
    Set sh = specswkb.Worksheets(sheetName)
    checkExists = (Not sh Is Nothing)
    On Error GoTo 0

    If Not checkExists Then ThrowError " Worksheet " & sheetName & " Not found!"
End Sub

'Errors
Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Exposed methods
Public Property Get ISpatialTables_Exists(ByVal varName As String) As Boolean
    ISpatialTables_Exists = Exists(varName)
End Property

Public Sub ISpatialTables_Add(ByVal formObject As IFormulas)
    Add formObject:=formObject
End Sub