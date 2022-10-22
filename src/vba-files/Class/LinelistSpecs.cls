VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "LinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Build")
'@ModuleDescription("Object containing specifications for the linelist")
'@IgnoreModule
Implements ILinelistSpecs

Option Explicit

Private Type TLinelistSpecs
  lldict As ILLdictionary
  llchoi As ILLchoice
End Type


Private this As TLinelistSpecs

Public Function Create(ByVal dict As ILLdictionary, ByVal choi As ILLchoice) As ILinelistSpecs

  With New LinelistSpecs
    Set .Dictionary = dict
    Set .Choices = choi
    Set Create = .Self
  End With
End Function



Public Property Get Self As ILinelistSpecs
  Set Self = Me
End Property

Public Property Get Dictionary() As ILLdictionary
  Set Dictionary = this.lldict
End Property

Public Property Get Choices() As ILLchoice
  Set Choices = this.llchoi
End Property

Public Property Set Dictionary(ByVal dict As ILLdictionary)
  Set this.lldict = dict
End Property

Public Property Set Choices(ByVal choi As ILLchoice)
  Set this.llchoi = choi
End Property


'Get Categories for one variable
Private Function Categories(ByVal varName As String) As BetterArray
  
  Dim categoryData As BetterArray
  Dim vars As ILLVariables
  Dim dict As ILLdictionary
  Dim choi As ILLChoice
  Dim categoryName As String

  Set dict = Dictionary()
  Set choi = Choices()
  Set categoryData = New BetterArray
  Set vars = LLVariables.Create(dict)

  categoryName = vars.Value(varName := varName, colName := "control details")
  Set categoryData = choi.Categories(categoryName)

  Set Categories = categoryData.Clone()

End Function



'Implement the interfaces

Public Property Get ILinelistSpecs_Dictionary() As ILLdictionary
  Set ILinelistSpecs_Dictionary = Dictionary()
End Property

Public Property Get ILinelistSpecs_Choices() As ILLchoice
  Set ILinelistSpecs_Choices = Choices()
End Property


'Categories of one variable
Public Function ILinelistSpecs_Categories(ByVal varName As String) As BetterArray
  Set ILinelistSpecs_Categories = Categories(varName)
End function