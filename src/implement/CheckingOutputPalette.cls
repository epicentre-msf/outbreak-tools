VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckingOutputPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Resolves typography and colour settings for CheckingOutput"

'@Folder("Implement")
'@ModuleDescription("Translates output level/colour tokens into font and fill settings")

Option Explicit

'@section State
'===============================================================================

Private Type TCheckingOutputPalette
    'No state required; palette is deterministic.
End Type

Private this As TCheckingOutputPalette

Private Const CLASSNAME As String = "CheckingOutputPalette"

Private Type TStyle
    FontSize As Long
    FontColour As Long
    FillColour As Long
    BorderColour As Long
    Bold As Boolean
End Type

'@section Public API
'===============================================================================

Public Function Resolve(ByVal level As Byte, ByVal colourName As String) As TStyle
    Dim style As TStyle

    style = BaseStyle(level)
    style = ApplyColour(style, LCase$(Trim$(colourName)))

    Resolve = style
End Function

'@section Internal helpers
'===============================================================================

Private Function BaseStyle(ByVal level As Byte) As TStyle
    Dim style As TStyle

    Select Case level
        Case 1
            style.FontSize = 20
            style.FontColour = RGB(0, 0, 128)
            style.FillColour = RGB(230, 235, 245)
            style.BorderColour = RGB(0, 0, 64)
            style.Bold = True
        Case 2
            style.FontSize = 14
            style.FontColour = RGB(0, 0, 128)
            style.FillColour = RGB(230, 235, 245)
            style.BorderColour = RGB(0, 0, 64)
            style.Bold = True
        Case Else
            style.FontSize = 10
            style.FontColour = RGB(0, 128, 64)
            style.FillColour = RGB(230, 255, 230)
            style.BorderColour = RGB(120, 200, 120)
            style.Bold = False
    End Select

    BaseStyle = style
End Function

Private Function ApplyColour(ByVal style As TStyle, ByVal colourKey As String) As TStyle
    Select Case colourKey
        Case "purple"
            style.FontColour = RGB(112, 48, 160)
            style.FillColour = RGB(244, 236, 255)
            style.BorderColour = RGB(203, 167, 226)
        Case "red"
            style.FontColour = RGB(192, 0, 0)
            style.FillColour = RGB(255, 235, 238)
            style.BorderColour = RGB(224, 102, 102)
        Case "orange"
            style.FontColour = RGB(167, 106, 0)
            style.FillColour = RGB(255, 240, 220)
            style.BorderColour = RGB(236, 184, 120)
        Case "grey"
            style.FontColour = RGB(96, 97, 100)
            style.FillColour = RGB(242, 243, 245)
            style.BorderColour = RGB(205, 206, 210)
        Case Else
            'Keep base style
    End Select

    ApplyColour = style
End Function

End Class
