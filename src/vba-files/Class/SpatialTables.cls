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
Private Const HF As String = "hf"
Private Const CONCAT As String = "concat_"

Public Function Create(ByVal tabl As ICrossTable) As ISpatialTables
    
    'output worksheet to add the spatial tables
    CheckRequirements tabl '(tabl is exists and spatial_tables__ sheet exists)

    With New SpatialTables
        Set .Table = tabl
        Set Create = .Self
    End With

End Function

Public Property Get Self() As ISpatialTables
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  'The worksheet is the sapatial table worksheet (where to write tables for admins)
  
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
  'Find the last column of the spatial table sheet.
  
    Dim sh As Worksheet
    Dim colIndex As Long
    Set sh = Wksh()
    With sh
        colIndex = .Cells(1, .Columns.Count).End(xlToLeft).Column
        colIndex = IIf(colIndex = 3, 7, colIndex + 2)
    End With
    LastColumn = colIndex
End Property


Private Property Get CurrentHeaderName(ByVal varName As String, ByVal adminName As String) As String
    CurrentHeaderName = "tabl_" & adminName & "_" & varName
End Property

Private Property Get CurrentLoName(ByVal varName As String, ByVal adminName As String) As String
  CurrentLoName = "spatial_" & adminName & "_" & varName
End Property

'Test if a variable exists in a sheet
Private Function Exists(ByVal varName As String) As Boolean
  Dim counter As Long
  Dim testExistence As Boolean
  Dim loName As String

  For counter = 1 To 4
    loName = CurrentLoName(varName, ADMIN & counter)
    testExistence = LoExists(loName)
    If Not testExistence Then Exit For
  Next

  Exists = testExistence
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

Private Sub AddVarNameToList(ByVal varName As String)
  Dim sh As Worksheet
  Dim Lo As ListObject
  Dim nbRows As Long

  Set sh = Wksh()
  Set Lo = sh.ListObjects("listofgeovars")

  'Add the variable to the listobject
  nbRows = Application.WorksheetFunction.CountA(Lo.Range)
  sh.Cells(nbRows + 1, Lo.Range.Column).Value = varName
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
    Dim finalFormula As String
    Dim adminName As String
    Dim rowVarName As String 'Variable name on rows for formulas
    Dim inputVarName As String
    Dim endCounter As Integer
    Dim spatType As String

    Set tabl = Table
    Set specs = tabl.Specifications
    spatType = specs.SpatialTableTypes()
    inputVarName = specs.Value("row")
    varName = inputVarName & "_" & specs.TableId()
    
    If Exists(varName) Then Exit Sub
    If Not LoExists("listofgeovars") Then Prepare
    
    'Add this variable name to the list of geovars (which contains the ids)
    AddVarNameToList varName
    
    Set sh = Wksh()
    
    With sh
      
      endCounter = Switch(spatType = "geo", 4, spatType = "hf", 1, True, 4)

      For counter = 1 To endCounter
        
        ClearTableConditions

        loName = Switch(spatType = "geo", CurrentLoName(varName, ADMIN & counter), _ 
                        spatType = "hf", CurrentLoName(varName, HF))

        colIndex = LastColumn()
        adminName = Switch(spatType = "geo", ADMIN & counter, spatType = "hf", HF)
        rowVarName = Switch(spatType = "geo", CONCAT & adminName & "_" & inputVarName, spatType = "hf", HF & "_" & inputVarName)
        'Add the table names for admin 1, only admin1, for admin2, admin1 and admin2, etc...
        headerName = CurrentHeaderName(varName, adminName)
        Set cellRng = .Cells(1, colIndex)
        cellRng.Value = headerName
        
        'CellRng IS  THE RANGE WITH THE HEADERNAME OF THE COLUMN OF CONCAT VALUES
        
        'Move to one row bellow to add conditions on rows
        AddTableConditions rowVarName, "= " & cellRng.Cells(2, 1).Address(Rowabsolute:=False)

        'Formula value for total
        cellRng.Cells(1, 2).Value = "formula_" & ADMIN & "_" & counter
        cellRng.Cells(1, 3).Value = "population_" & ADMIN & "_" & counter
        
        'final formula
        finalFormula = CurrentExcelFormula(formObject)

        'Test final formula to see if it returns something
        If finalFormula <> vbNullString Then cellRng.Cells(2, 2).FormulaArray = "= " & finalFormula
        
        'Formula for the population of the admin level
        finalFormula = "= GeoPopulation(" & counter - 1 & ", " & cellRng.Cells(2, 1).Address(Rowabsolute:=False) & ")"
        cellRng.Cells(2, 3).Formula = finalFormula

        'Define the listObject
        Set rng = .Range(cellRng.Cells(1, 1), cellRng.Cells(2, 3))

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
