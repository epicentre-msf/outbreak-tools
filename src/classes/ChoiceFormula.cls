VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChoiceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "choice formula"


'@Folder("Dictionary")
'@ModuleDescription("A class for choices formula")
'@IgnoreModule

Option Explicit
Implements IChoiceFormula

Private Type TChoiceFormula
  formula As String
End Type


'Exposed methods
Private This As TChoiceFormula

Public Function Create(ByVal formula As String) As IChoiceFormula
  With New ChoiceFormula
      .ChoiceSpecifications = formula
      Set Create = .Self
  End With
End Function


Public Property Get Self() As IChoiceFormula
  Set Self = Me
End Property

Public Property Get ChoiceSpecifications() As String
  ChoiceSpecifications = This.formula
End Property

Public Property Let ChoiceSpecifications(ByVal formula As String)
  This.formula = Application.WorksheetFunction.Trim(formula)
End Property

'Test the validity of the choice formula
Private Property Get Valid() As Boolean
  Dim formula As String
  formula = ChoiceSpecifications
  Valid = (InStr(1, formula, "CHOICE_FORMULA") = 1) And (Right(formula, 1) = ")")
End Property

Private Property Get casewhenObject() As ICaseWhen
  Dim formula As String
  Dim convertedFormula As String 'converted formula to case_when.

  formula = ChoiceSpecifications()
  'convert the choice_formula to casewhen for processing
  'The comma here is important!
  convertedFormula = Application.WorksheetFunction.Trim(Replace(formula, Split(formula, ",")(0) & ",", ""))
  convertedFormula = "CASE_WHEN(" & convertedFormula
  convertedFormula = Left(convertedFormula, Len(convertedFormula) - 1)
  convertedFormula = convertedFormula & ", " & chr(34) & chr(34) & ")"

  Set casewhenObject = CaseWhen.Create(convertedFormula)
  
End Property

'Get the parsed formula of the CHOICE_FORMULA
Private Function parsedFormula() As String
  Dim casewhenObj As ICaseWhen
  If Valid() Then
    Set casewhenObj = casewhenObject()
    parsedFormula = casewhenObj.parsedFormula()
  End If
End Function

'Retrieve the categories of only the mentionned values
Private Function Categories() As BetterArray
  
  Dim categoriesTable As BetterArray
  Dim casewhenObj As ICaseWhen

  Set categoriesTable = New BetterArray
  
  If Valid() Then
    Set casewhenObj = casewhenObject()
    Set categoriesTable = casewhenObj.Categories()
  End If

  Set Categories = categoriesTable.Clone()
End Function


'Get the choice name from the choice formula
Private Function choiceName() As String
  Dim formula As String
  If Valid() Then
      formula = ChoiceSpecifications()
      choiceName = Replace(Split(formula, ",")(0), "CHOICE_FORMULA(", "")
  End If
End Function


'Interface
Public Property Get IChoiceFormula_parsedFormula() As String
  IChoiceFormula_parsedFormula = parsedFormula()
End Property

'List of categories
Public Property Get IChoiceFormula_Categories() As BetterArray
  Set IChoiceFormula_Categories = Categories()
End Property

'Validity test
Public Property Get IChoiceFormula_Valid() As Boolean
  IChoiceFormula_Valid = Valid()
End Property

Public Property Get IChoiceFormula_ChoiceName() As String
  IChoiceFormula_ChoiceName = choiceName()
End Property
