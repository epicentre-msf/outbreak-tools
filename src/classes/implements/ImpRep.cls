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

Private Type TRep
  impsh As Worksheet
  impForm As Object
End Type

Private this As TImp
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
    Set .FormObject = formObj
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

Public Property Get FormObject() As Object
  Set FormObject = this.impForm
End Property

Public Property Set FormObject(ByVal fobj As Object)
  Set this.impForm = fobj
End Property

'Clear the report
Private Sub Clear()
  Dim sh As Worksheet
  Dim Lo As ListObject
  Set sh = Wksh()
  For Each Lo in sh.ListObjects
    'DatabodyRange could be nothing resulting in Error
    On Error Resume Next 
     Lo.DataBodyRange.Delete
    On Error GoTo 0
  Next
End Sub


'Add elements to report at sheet levels --------
'scope could take 2 values: 1 - sheet not imported, 2, sheet not touched

'Add sheets not imported or not touched to the import report sheet
Private Sub AddSheetsToReport(ByVal scope As Byte, ByVal sheetName As String)
  Dim sh As Worksheet
  Dim loName As String 
  Dim Lo As ListObject
  Dim rng As Range
  Dim endRow As Long

  Set sh = Wksh()
  loName = Switch(scope = 1, TABSHEETNOTIMP, scope = 2, TABSHEETNOTTOU)
  
  Set Lo = sh.ListObjects(loName)

  'Where to paste the data
  If Lo.DataBodyRange Is Nothing Then 
    endRow = 2
  Else
    endRow = Lo.Range.Rows.Count + 1  
  End If

  Lo.Range.Cells(endRow, 1).
  
  sheetsList.ToExcelRange Lo.Range.Cells(endRow, 1)
  'resize the listObject
  Set rng = sh.Range(Lo.Range.Cells(1, 1), Lo.Range.Cells(sheetsList.Length + endRow - 1, 1))
  Lo.Resize rng
End Sub

'Test if the import report sheet contains sheet not imported or not touched
Private Property Get ContainsSheet(ByVal scope As Byte,  _ 
                                  ByVal sheetName As String) As Boolean

  Dim rng As Range
  Dim sh As Worksheet
  Dim loName As String
  Dim contains As Boolean

  contains = False

  Set sh = Wksh()
  loName = Switch(scope = 1, TABSHEETNOTIMP, scope = 2, TABSHEETNOTTOU)
  Set rng = sh.ListObjects(loName).DataBodyRange

  If Not (rng Is Nothing) Then
    contains = Not (rng.Find(What:=sheetName, LookAt:=xlWhole, MatchCase:=True) Is Nothing)
  End If

  ContainsSheet = contains

End Property


'Variables imported/not touched
'scope takes 2 values 1: variables not imported, 2: variable in current workbook not found
'in imported workbook

Private Sub AddVariablesToReport(ByVal scope As Byte, ByVal sheetName As String, _ 
                                 ByVal varsList As BetterArray)
  Dim sh As Worksheet
  Dim loName As String 
  Dim Lo As ListObject
  Dim rng As Range
  Dim endRow As Long

  Set sh = Wksh()
  loName = Switch(scope = 1, TABVARSNOTIMP, scope = 2, TABVARSNOTTOU)
  Set Lo = sh.ListObjects(loName)

  If Lo.DataBodyRange Is Nothing Then
    endRow = 2
  Else
    endRow = Lo.Range.Rows.Count + 1
  End If

  'write the list of the variables
  'The first column is for variables, the second one for the corresponding sheets
  varsList.ToExcelRange Lo.Range.Cells(endRow, 1)

  'Add the sheetNames
  For counter = endRow To varsList.Length + endRow - 1
    Lo.Range.Cells(endRow, 2).Value = sheetName
  Next

  'Resize the listObject
  Set rng = sh.Range(Lo.Range.Cells(1, 1), Lo.Range.Cells(varsList.Length + endRow - 1, 2))
  Lo.Resize rng    
End Sub

Private Property Get ContainsVariable(ByVal scope As Byte, ByVal sheetName As String, _ 
                                      ByVal varName As String) As Boolean
  
  Dim rng As Range
  Dim sh As Worksheet
  Dim loName As String
  Dim contains As Boolean
  Dim cellRng As Range

  contains = False
  loName = Switch(scope = 1, TABVARSNOTIMP, scope = 2, TABVARSNOTTOU)
  
  'The second column is for worksheets
  Set rng = sh.ListObjects(loName).ListColumns(2).DataBodyRange
  
  'Test if the sheet name exists before testing variable
  If Not rng Is Nothing Then
    If Not (rng.Find(What:=sheetName, LookAt:=xlWhole, MatchCase:=True) Is Nothing) Then
      Set cellRng = rng.Find(What:=sheetName, LookAt:=xlWhole, MatchCase:=True)
      Do While (cellRng.Value = sheetName)

        'The first column contains the variable names, so I will test at the left
        If cellRng.Offset(, -1).Value = varName Then
          contains = True
          Exit Do
        End If

        Set cellRng = cellRng.Offset(1) 
      Loop
    End If      
  End If

  ContainsVariable = contains
End Property

'Show the report
Private Sub Show()
  Dim sh As Worksheet
  Dim fobj As Object
  Dim rng As Range
  Dim Lo As ListObject

  Set fobj = FormObject()

  'Sheets Not imported
  Set sh = Wksh()

  Set Lo = sh.ListObjects(TABSHEETNOTIMP)
  If Not Lo.DataBodyRange Is Nothing Then
    Set rng = Lo.DataBodyRange
    fobj.LST_ImpRepSheet.ColumnCount = 1
    fobj.LST_ImpRepSheet.List = rng.Value
  End If

  'Sheets Not touched
  Set Lo =  sh.ListObjects(TABSHEETNOTTOU)
  If Not Lo.DataBodyRange Is Nothing Then
    Set rng = Lo.DataBodyRange
    fobj.LST_ImpLLSheet.ColumnCount = 1
    fobj.LST_ImpLLSheet.List = rng.Value
  End If

  'Variables Not Imported
  Set Lo = sh.ListObjects(TABVARSNOTIMP)
  If Not Lo.DataBodyRange Is Nothing Then
    Set rng = Lo.DataBodyRange
    fobj.LST_ImpRepVarImp.ColumnCount = 2
    fobj.LST_ImpRepVarImp.List = rng.Value
  End If

  'Variable Not touched
  Set Lo = sh.ListObjects(TABVARSNOTTOU)
  If Not Lo.DataBodyRange Is Nothing Then
    Set rng = Lo.DataBodyRange
    fobj.LST_ImpRepVarLL.ColumnCount = 2
    fobj.LST_ImpRepVarLL.List = rng.Value
  End If

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
Private Sub LoExists(ByVal sh As WorkSheet, ByVal loName As String)
  Dim Lo As ListObject                         
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
  Show
End Sub

Private Sub IImpRep_AddSheetsToReport(ByVal scope As Byte, ByVal sheetsList As Variant)
  AddSheetsToReport scope, sheetsList
End Sub

Private Sub IImpRep_AddVariablesToReport(ByVal scope As Byte, ByVal sheetName As String, ByVal varsList As BetterArray)
  AddVariablesToReport scope, sheetName, varsList
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