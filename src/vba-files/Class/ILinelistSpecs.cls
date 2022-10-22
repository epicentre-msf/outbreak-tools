VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ILinelistSpecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = ""

'@Folder("Build")
'@ModuleDescription("Interface of Linelist specification")
'@IgnoreModule

Option Explicit

'Exposed methods


Public Property Get Dictionary() As ILLdictionary :  End Property
Public Property Get Choices() As ILLChoice :  End Property






'Exposed methods
'Categories of one variable
Public Function Categories(Byval varName As String) As BetterArray : End function