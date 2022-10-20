VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Data Used for the Formula Object"


'@Folder("Formula")
'@ModuleDescription("Formula Data Class")
'@IgnoreModule

Private Const CLASSNAME As String = "FormulaData"

Option Explicit
Implements IFormulaData

Private Type TFormulaData
  fWksh As Worksheet
  fSpecRange As Range
  fFormRange As Range
End Type

Private this As TFormulaData


'instanciation
Public Function Create(ByVal formWksh As Worksheet, ByVal excelFormLo As String, ByVal specCharLo As String) As IFormulaData

  If (formWksh Is Nothing) Then ThrowErrorUnableInstanciate "No Worksheet"
  If (formWksh.ListObjects(specCharLo) Is Nothing) Then ThrowErrorUnableInstanciate "No special character listobject"
  If (formWksh.ListObjects(specCharLo).DataBodyRange Is Nothing Or _
   formWksh.ListObjects(excelFormLo).DataBodyRange Is Nothing) Then
    ThrowErrorUnableInstanciate "No data in formulas or special characters"
  End If
  If (formWksh.ListObjects(excelFormLo) Is Nothing) Then ThrowErrorUnableInstanciate "No formula listobject"
  If (formWksh.ListObjects(excelFormLo).ListColumns("ENG") Is Nothing) Then ThrowErrorUnableInstanciate "No english column"
  If (formWksh.ListObjects(specCharLo).ListColumns(2) Is Nothing) Then ThrowErrorUnableInstanciate "No special characters column"

  With New FormulaData
      Set .Wksh = formWksh
      Set .FormulaRange = formWksh.ListObjects(excelFormLo).ListColumns("ENG").DataBodyRange
      Set .SpecialCharactersRange = formWksh.ListObjects(specCharLo).ListColumns(2).DataBodyRange
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
  Set FormulaRange = this.fFormRange
End Property

Public Property Set FormulaRange(ByVal rng As Range)
  Set this.fFormRange = rng
End Property

Public Property Get SpecialCharactersRange() As Range
  Set SpecialCharactersRange = this.fSpecRange
End Property

Public Property Set SpecialCharactersRange(ByVal rng As Range)
    Set this.fSpecRange = rng
End Property

'Methods
Private Function SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
  Dim rng As Range

  Set rng = SpecialCharactersRange
  SpecialCharacterIncludes = (Not rng.Find(What:=specialChar, LookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function

Private Function ExcelFormulasIncludes(ByVal excelFormula As String) As Boolean
  Dim rng As Range

  Set rng = FormulaRange
  ExcelFormulasIncludes = (Not rng.Find(What:=excelFormula, LookAt:=xlWhole, MatchCase:=False) Is Nothing)
End Function


'Errors ---------------------------------------------------------------------------------
Private Sub ThrowErrorUnableInstanciate(ByVal errorMessage As String)
  Err.Raise VBA.vbObjectError + 515, CLASSNAME, "Unable to create an instance of the class : " & errorMessage
End Sub

'Exposed methods
Public Function IFormulaData_SpecialCharacterIncludes(ByVal specialChar As String) As Boolean
  IFormulaData_SpecialCharacterIncludes = SpecialCharacterIncludes(specialChar)
End Function

Public Function IFormulaData_ExcelFormulasIncludes(ByVal excelFormula As String) As Boolean
  IFormulaData_ExcelFormulasIncludes = ExcelFormulasIncludes(excelFormula)
End Function

