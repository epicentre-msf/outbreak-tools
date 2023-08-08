VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImpRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Import Reports"


'@PredeclaredId
'@Folder("Linelist")
'@ModuleDescription("Import Reports")

Option Explicit
Implements IImpRep

Private Type TRep
  impsh As Worksheet
  impForm As Object
  sorted As Boolean
End Type

Private this As TRep
Private Const CLASSNAME As String = "ImpRep"

'ListObjects to have in the import sheet
Private Const TABSHEETNOTIMP As String = "reptab_sheetsNotImp"
Private Const TABSHEETNOTTOU As String = "reptab_sheetsNotTouch"
Private Const TABSVARSNOTIMP As String = "reptab_varsNotImp"
Private Const TABSVARSNOTTOU As String = "reptab_varsNotTouch"


'Exposed methods
Public Function Create(ByVal formObj As Object, impsh As Worksheet) As IImpRep
  
  CheckRequirements formObj, impsh
  
  With New ImpRep
    Set .Wksh = impsh
    Set .formObject = formObj
    Set Create = .Self()
  End With

End Function

Public Property Get Self() As IImpRep
  Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
  Set Wksh = this.impsh
End Property

Public Property Set Wksh(ByVal impsh As Worksheet)
  Set this.impsh = impsh
End Property

Public Property Get formObject() As Object
  Set formObject = this.impForm
End Property

Public Property Set formObject(ByVal fobj As Object)
  Set this.impForm = fobj
End Property

'Clear the report
Private Sub Clear()
  Dim sh As Worksheet
  Dim Lo As listObject
  Set sh = Wksh()
  For Each Lo In sh.ListObjects
    'DatabodyRange could be nothing resulting in Error
    If Not (Lo.DataBodyRange Is Nothing) Then Lo.DataBodyRange.Delete
  Next
End Sub

Private Sub SetSortedToFalse()
  this.sorted = False
End Sub

Private Sub SetSortedToTrue()
  this.sorted = True
End Sub

Private Property Get IsSorted() As Boolean
  IsSorted = this.sorted
End Property

'Add elements to report at sheet levels --------
'scope could take 2 values: 1 - sheet not imported, 2, sheet not touched

'Add sheets not imported or not touched to the import report sheet
Private Sub AddSheetsToReport(ByVal scope As Byte, ByVal sheetName As String)
  Dim sh As Worksheet
  Dim loName As String
  Dim Lo As listObject
  Dim rng As Range
  Dim EndRow As Long

  Set sh = Wksh()
  loName = Switch(scope = 1, TABSHEETNOTIMP, scope = 2, TABSHEETNOTTOU)
  
  Set Lo = sh.ListObjects(loName)

  'Where to paste the data
  If Lo.DataBodyRange Is Nothing Then
    EndRow = 2
  Else
    EndRow = Lo.Range.Rows.Count + 1
  End If

  Lo.Range.Cells(EndRow, 1).Value = sheetName
  
  'resize the listObject
  Set rng = sh.Range(Lo.Range.Cells(1, 1), Lo.Range.Cells(EndRow, 1))
  Lo.Resize rng

  SetSortedToFalse
End Sub

'Test if the import report sheet contains sheet not imported or not touched
Private Property Get ContainsSheet(ByVal scope As Byte, _
                                  ByVal sheetName As String) As Boolean

  Dim rng As Range
  Dim sh As Worksheet
  Dim loName As String
  Dim Contains As Boolean

  Contains = False

  Set sh = Wksh()
  loName = Switch(scope = 1, TABSHEETNOTIMP, scope = 2, TABSHEETNOTTOU)
  Set rng = sh.ListObjects(loName).DataBodyRange

  If Not (rng Is Nothing) Then
    Contains = Not (rng.Find(What:=sheetName, lookAt:=xlWhole, MatchCase:=True) Is Nothing)
  End If

  ContainsSheet = Contains
End Property


'Variables imported/not touched
'scope takes 2 values 1: variables not imported, 2: variable in current workbook not found
'in imported workbook

'Add only one variable to the report sheet
Private Sub AddVariablesToReport(ByVal scope As Byte, ByVal sheetName As String, _
                                 ByVal varName As String)
  Dim sh As Worksheet
  Dim loName As String
  Dim Lo As listObject
  Dim rng As Range
  Dim EndRow As Long

  Set sh = Wksh()
  loName = Switch(scope = 1, TABSVARSNOTIMP, scope = 2, TABSVARSNOTTOU)
  Set Lo = sh.ListObjects(loName)

  If Lo.DataBodyRange Is Nothing Then
    EndRow = 2
  Else
    EndRow = Lo.Range.Rows.Count + 1
  End If

  'write the variables and the sheetName
  Lo.Range.Cells(EndRow, 1).Value = varName
  Lo.Range.Cells(EndRow, 2).Value = sheetName

  'Resize the listObject
  Set rng = sh.Range(Lo.Range.Cells(1, 1), Lo.Range.Cells(EndRow, 2))
  Lo.Resize rng

  SetSortedToFalse
End Sub

'Add a list of variables to the report sheet
Private Sub AddVariablesListToReport(ByVal scope As Byte, ByVal sheetName As String, _
                                     ByVal varsList As BetterArray)
  Dim sh As Worksheet
  Dim loName As String
  Dim Lo As listObject
  Dim rng As Range
  Dim EndRow As Long
  Dim counter As Long

  Set sh = Wksh()
  loName = Switch(scope = 1, TABSVARSNOTIMP, scope = 2, TABSVARSNOTTOU)
  
  Set Lo = sh.ListObjects(loName)

  If Lo.DataBodyRange Is Nothing Then
    EndRow = 2
  Else
    EndRow = Lo.Range.Rows.Count + 1
  End If

  'write the list of the variables
  'The first column is for variables, the second one for the corresponding sheets
  varsList.ToExcelRange Lo.Range.Cells(EndRow, 1)

  'Add the sheetNames
  For counter = EndRow To varsList.Length + EndRow - 1
    Lo.Range.Cells(counter, 2).Value = sheetName
  Next

  'Resize the listObject

  Set rng = sh.Range(Lo.Range.Cells(1, 1), Lo.Range.Cells(varsList.Length + EndRow - 1, 2))
  Lo.Resize rng

  SetSortedToFalse
End Sub

'Test if the report contains a variable
Private Property Get ContainsVariable(ByVal scope As Byte, ByVal sheetName As String, _
                                      ByVal varName As String) As Boolean
  
  Dim rng As Range
  Dim sh As Worksheet
  Dim loName As String
  Dim Contains As Boolean
  Dim cellRng As Range
  Dim sh As Worksheet

  'Important to sort the worksheet before.
  Sort
  Set sh = Wksh()

  Contains = False
  loName = Switch(scope = 1, TABSVARSNOTIMP, scope = 2, TABSVARSNOTTOU)

  'The second column is for worksheets
  Set rng = sh.ListObjects(loName).ListColumns(2).DataBodyRange
  
  'Test if the sheet name exists before testing variable
  If Not rng Is Nothing Then
    On Error Resume Next
    Set cellRng = rng.Find(What:=sheetName, lookAt:=xlWhole, MatchCase:=True)
    On Error GoTo 0

    If Not (cellRng Is Nothing) Then
      Do While (cellRng.Value = sheetName)

        'The first column contains the variable names, so I will test at the left
        If cellRng.Offset(, -1).Value = varName Then
          Contains = True
          Exit Do
        End If

        Set cellRng = cellRng.Offset(1)
      Loop
    End If
  End If

  ContainsVariable = Contains
End Property

'Sort the listObjects on the different columns
Private Sub Sort()
  Dim rng As Range
  Dim Lo As listObject
  Dim sh As Worksheet

  Set sh = Wksh()
  'Sort sheets not imported
  
  If IsSorted() Then Exit Sub

  Set Lo = sh.ListObjects(TABSHEETNOTIMP)
  Set rng = Lo.Range
  On Error Resume Next
  rng.Sort key1:=rng, Order1:=xlDescending, Header:=xlYes
  On Error GoTo 0

  'Sort sheets not touched
  Set Lo = sh.ListObjects(TABSHEETNOTTOU)
  Set rng = Lo.Range
  On Error Resume Next
  rng.Sort key1:=rng, Order1:=xlDescending, Header:=xlYes
  On Error GoTo 0

  'Sort variables not imported
  Set Lo = sh.ListObjects(TABSVARSNOTIMP)
  Set rng = Lo.Range
  On Error Resume Next
  rng.Sort key1:=Lo.ListObjects(2).Range, Order1:=xlDescending, Header:=xlYes
  On Error GoTo 0

  'Sort variables not touched
  Set Lo = sh.ListObjects(TABSVARSNOTTOU)
  Set rng = Lo.Range
  On Error Resume Next
  rng.Sort key1:=Lo.ListObjects(2).Range, Order1:=xlDescending, Header:=xlYes
  On Error GoTo 0

  'Update sorted boolean
  SetSortedToTrue
End Sub

'Show the report
Private Sub Show()
  Dim sh As Worksheet
  Dim fobj As Object
  Dim rng As Range
  Dim Lo As listObject

  Set fobj = formObject()

  'Sheets Not imported
  Set sh = Wksh()

  Set Lo = sh.ListObjects(TABSHEETNOTIMP)
  Set rng = Lo.Range.Offset(1)
  fobj.LST_ImpRepSheet.ColumnCount = 1
  fobj.LST_ImpRepSheet.List = rng.Value

  'Sheets Not touched
  Set Lo = sh.ListObjects(TABSHEETNOTTOU)
  Set rng = Lo.Range.Offset(1)
  fobj.LST_ImpLLSheet.ColumnCount = 1
  fobj.LST_ImpLLSheet.List = rng.Value

  'Variables Not Imported
  Set Lo = sh.ListObjects(TABSVARSNOTIMP)
  Set rng = Lo.Range.Offset(1)
  fobj.LST_ImpRepVarImp.ColumnCount = 2
  fobj.LST_ImpRepVarImp.List = rng.Value

  'Variable Not touched
  Set Lo = sh.ListObjects(TABSVARSNOTTOU)
  Set rng = Lo.Range.Offset(1)
  fobj.LST_ImpRepVarLL.ColumnCount = 2
  fobj.LST_ImpRepVarLL.List = rng.Value

  fobj.Show
End Sub

  'Errors checking at instanciation
Private Sub CheckRequirements(ByVal formObj As Object, ByVal impsh As Worksheet)
  Dim counter As Long
  Dim loName As String
  Dim loArray As BetterArray
  If formObj Is Nothing Then ThrowErrorUnableInstanciate "Form object not defined"
  If TypeName(formObj) <> "F_ImportRep" Then ThrowErrorUnableInstanciate "The provided form is not F_ImportRep"
  If impsh Is Nothing Then ThrowErrorUnableInstanciate "Report sheet not defined"
  'Test if the listObjects for report are present
  Set loArray = New BetterArray
  loArray.Push TABSHEETNOTIMP, TABSHEETNOTTOU, TABSVARSNOTIMP, TABSVARSNOTTOU
  For counter = loArray.LowerBound To loArray.UpperBound
    loName = loArray.Item(counter)
    LoExists impsh, loName
  Next
End Sub
  
'Returns true or false if a sheet or a range exists
Private Sub LoExists(ByVal sh As Worksheet, ByVal loName As String)
  Dim Lo As listObject
  Dim checkExists As Boolean

  On Error Resume Next
  Set Lo = sh.ListObjects(loName)
  checkExists = (Not Lo Is Nothing)
  On Error GoTo 0

  If Not checkExists Then ThrowErrorUnableInstanciate " ListObject " & loName & " Not found!"
End Sub
 
  'Errors managements
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, _
   "Unable to create an instance of the class " & CLASSNAME & ": " & errorMessage
End Sub

Private Sub ThrowError(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 516, CLASSNAME, "Error : " & errorMessage
End Sub

'Iterface
Private Sub IImpRep_Show()
  Sort
  Show
End Sub

Private Sub IImpRep_AddSheetsToReport(ByVal scope As Byte, ByVal sheetName As String)
  AddSheetsToReport scope, sheetName
End Sub

Private Sub IImpRep_AddVariablesToReport(ByVal scope As Byte, ByVal sheetName As String, ByVal varName As String)
  AddVariablesToReport scope, sheetName, varName
End Sub

Private Sub IImpRep_AddVariablesListToReport(ByVal scope As Byte, ByVal sheetName As String, ByVal varsList As BetterArray)
  AddVariablesListToReport scope, sheetName, varsList
End Sub

Private Sub IImpRep_Clear()
  Clear
End Sub

Private Property Get IImpRep_ContainsSheet(ByVal scope As Byte, ByVal sheetName As String) As Boolean
  IImpRep_ContainsSheet = ContainsSheet(scope, sheetName)
End Property

Private Property Get IImpRep_ContainsVariable(ByVal scope As Byte, _
                                              ByVal sheetName As String, _
                                              ByVal varName As String) As Boolean

  IImpRep_ContainsVariable = ContainsVariable(scope, sheetName, varName)
End Property
