VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface of Linelist Buttons"

'@IgnoreModule UseMeaningfulName, HungarianNotation
'@Folder("Linelist")
'@ModuleDescription("Interface of Linelist Buttons")
'@Interface

Option Explicit

Public Enum ButtonScope
    ButtonScopeLarge = 1
    ButtonScopeSmall = 2
    ButtonScopeGeo = 3
End Enum

'Exposed methods
Public Sub Add(Optional ByVal actionCommand As String = vbNullString, Optional ByVal shapeLabel As String = vbNullString): End Sub
Public Sub Format(Optional ByVal fontColor As Variant = vbNullString, _
                  Optional ByVal interiorColor As Variant = vbNullString, _
                  Optional ByVal isBold As Boolean = False, _
                  Optional ByVal FontSize As Double = 0)
End Sub
