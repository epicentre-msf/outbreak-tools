VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormulaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A Class for dealing with conditions in formulas"
'@Folder("Formulas")
'@ModuleDescription("Conditons Used to create formulas for analysis")
'@IgnoreModule

Option Explicit
Implements IFormulaCondition

'Types of the class

Private Type TFormulaCondition
    Wkb As Workbook
    Conditions As BetterArray
    Variables As BetterArray
End Type

'Instanciation
Public Function Create(ByVal Wkb As Workbook, Conditions As BetterArray, Variables As BetterArray) As IFormulaCondition
    With .New FormulaCondition
        Set .Wkb = Wkb
        Set .Conditions = Conditions
        Set .Variables = Variables
        Set Create = .Self
    End With
End Function


Public Property Get Self() As IFormulaCondition
    Set Self = Me
End Property

'Wkb
Public Property Get Wkb() As Workbook
    Set Wkb = this.Wkb
End Property

Public Property Let Wkb(Byval Wkb as Workbook)
    Set this.Wkb = Wkb
End Property

'Conditions
Public Property Get Conditions() As BetterArray
    Set Conditions = this.Conditions
End Property


Public Property Let Self() As IFormulaCondition

End Property


Public Property Get Self() As IFormulaCondition

End Property

Public Property Let Self() As IFormulaCondition

End Property