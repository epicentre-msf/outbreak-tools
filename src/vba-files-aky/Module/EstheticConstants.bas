Attribute VB_Name = "EstheticConstants"
Option Private Module

'Constants used when defining esthetic things for the linelist and the
'Designer, as well as some function for esthetic work.
'We can tweak here the appearance of the linelist


'Fonts sizes
Public Const C_iLLShapesFont                As Integer = 9 'Fonts used when creating the shapes in the Linelist
Public Const C_iLLSheetFontSize             As Integer = 9 'Default font size on a linelist sheet
Public Const C_iLLSubSecFontSize            As Integer = 10 'Font size of the sub label in one sheet of type linelist
Public Const C_iLLMainSecFontSize           As Integer = 10 'Font size of the main label in one sheet of type linelist
Public Const C_iAdmSheetFontSize            As Integer = 14
Public Const C_iAdmTitleFontSize            As Integer = 20
Public Const C_iAnalysisFontSize            As Integer = 11

'Command buttons
Public Const C_iCmdWidth                    As Integer = 105 'Width of command added  on one sheet
Public Const C_iCmdHeight                   As Integer = 30 'Height of command added on one sheet

'Table Styles
Public Const C_sLLTableStyle                As String = "TableStyleLight16" 'Default table style of a linelist object

'COLORS   =========================================================================================================================================================================




'INTEGERS CONSTANTS ===============================================================================================================================================================

Public Const C_iLLButtonsRowHeight          As Integer = 30 'Row height of the first two rows on one sheet of type Linelist for buttons
Public Const C_iNumberOfBars                As Integer = 40 'Number of bars in the Progressbar in the designer
Public Const C_iLLSplitColumn               As Integer = 2 'Where to split columns of type linelist
Public Const C_iLLFirstColumnsWidth         As Integer = 22
Public Const C_iAdmLinesColor               As Integer = 32
Public Const C_iLLVarLabelHeight            As Integer = 80 'Row height of variable label on linelist sheet

'Number of lines to display for a time series table
Public Const C_iNbTime                      As Integer = 52
