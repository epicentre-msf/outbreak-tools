Attribute VB_Name = "FormLogicShowHidePrint"
Attribute VB_Description = "Events show/hide in the printed linelist"

'@IgnoreModule UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Events show/hide in the printed linelist")

Option Explicit

Private Sub LST_PrintNames_Click()
    ClickListShowHide (LST_PrintNames.ListIndex)
End Sub

Private Sub OPT_PrintShowHoriz_Click()
    ClickOptionsShowHide (LST_PrintNames.ListIndex)
End Sub

Private Sub OPT_PrintShowVerti_Click()
    ClickOptionsShowHide (LST_PrintNames.ListIndex)
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHide (LST_PrintNames.ListIndex)
End Sub

Private Sub CMD_PrintBack_Click()
    Me.Hide
End Sub

'Print the worksheet
Private Sub CMD_PrintLL_Click()
    Dim sh As Worksheet

    'Set up the sheet with some print Characteristics
    Set sh = ActiveSheet

    On Error Resume Next
    Application.PrintCommunication = False
    'Avoid printing rows and column number'
    With sh.PageSetup
        'Specifies the margins
        .LeftMargin = Application.InchesToPoints(0.04)
        .RightMargin = Application.InchesToPoints(0.04)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.20)
        .HeaderMargin = Application.InchesToPoints(0.31)
        .FooterMargin = Application.InchesToPoints(0.31)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintTitleRows = "$5:$8" 'Those are rows to always keep on title
        .PrintTitleColumns = ""
        .PrintComments = xlPrintNoComments
        .PrintNotes = False
        'The quality of the print
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        'Landscape and paper size
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .FirstPageNumber = xlAutomatic
        .ORDER = xlDownThenOver
        .BlackAndWhite = False
        'Print the whole area and fit all columns in the worksheet
        .Zoom = 90
        .FitToPagesWide = 1
        .FitToPagesTall = False
        'Print Errors to blanks
        .PrintArea = sh.ListObjects(1).Range.Address
        .PrintErrors = xlPrintErrorsBlank
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Me.Hide
    sh.PrintPreview
End Sub

Private Sub CMD_ColWidth_Click()
    ClickColWidth (LST_PrintNames.ListIndex)
End Sub