VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Required data to convert formulas to excel"

'@PredeclaredId
'@Folder("Formulas")
'@ModuleDescription("Required data to convert formulas to excel")

Private Const CLASSNAME As String = "FormulaData"
Private Const SPECCHAR As String = "T_ascii"
Private Const excelForm As String = "T_XlsFonctions"

Option Explicit
Implements IFormulaData

Private Type TFormulaData
    fWksh As Worksheet
End Type

Private this As TFormulaData

'instanciation
Public Function Create(ByVal formWksh As Worksheet) As IFormulaData
    CheckRequirements formWksh
    With New FormulaData
        Set .Wksh = formWksh
        Set Create = .Self
    End With
End Function

'Properties
Public Property Get Self() As IFormulaData
    Set Self = Me
End Property

Public Property Get Wksh() As Worksheet
    Set Wksh = this.fWksh
End Property

Public Property Set Wksh(ByVal formWksh As Worksheet)
    Set this.fWksh = formWksh
End Property

Public Property Get FormulaRange() As Range
    Set FormulaRange = Wksh.ListObjects(excelForm).ListColumns("ENG").DataBodyRange
End Property

Public Property Get SpecialCharactersRange() As Range
    Set SpecialCharactersRange = Wksh.ListObjects(SPECCHAR).ListColumns(2).DataBodyRange
End Property

'Methods
Private Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    Dim rng As Range

    Set rng = SpecialCharactersRange
    SpecialCharacterIncludes = (Not rng.Find(What:=specialChar, lookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

Private Function ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    Dim rng As Range

    Set rng = FormulaRange
    ExcelFormulasIncludes = (Not rng.Find(What:=ExcelFormula, lookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

'Errors management ---------------------------------------------------------------------------------

Private Sub CheckRequirements(ByVal sh As Worksheet)
    Dim Lo As ListObject
    Dim rng As Range

    If sh Is Nothing Then ThrowErrorUnableInstanciate "No worksheet found"

    On Error Resume Next
    Set Lo = sh.ListObjects(excelForm)
    On Error GoTo 0

    'Special characters listObject
    If Lo Is Nothing Then ThrowErrorUnableInstanciate "No special character listObject"
    
    On Error Resume Next
    Set rng = Lo.ListColumns("ENG").DataBodyRange
    On Error GoTo 0

    If rng Is Nothing Then ThrowErrorUnableInstanciate "There is no column ENG in table " & Lo.Name

    'Excel formulas listObject
    Set rng = Nothing
    Set Lo = Nothing

    On Error Resume Next
    Set Lo = sh.ListObjects(SPECCHAR)
    On Error GoTo 0

    If Lo Is Nothing Then ThrowErrorUnableInstanciate "No Excel formula ListObject"
    
    On Error Resume Next
    Set rng = Lo.ListColumns(2).DataBodyRange
    On Error GoTo 0

    If rng Is Nothing Then ThrowErrorUnableInstanciate "There is no column for ascii characters in table " & Lo.Name
End Sub

Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
    Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Exposed methods
Private Function IFormulaData_SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
    IFormulaData_SpecialCharacterIncludes = SpecialCharacterIncludes(specialChar)
End Function

Private Function IFormulaData_ExcelFormulasIncludes(ByVal ExcelFormula As String) As Boolean
    IFormulaData_ExcelFormulasIncludes = ExcelFormulasIncludes(ExcelFormula)
End Function

