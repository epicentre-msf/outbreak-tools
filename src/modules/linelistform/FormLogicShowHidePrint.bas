Attribute VB_Name = "FormLogicShowHidePrint"
Attribute VB_Description = "Form code-behind for F_ShowHidePrint"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHidePrint")

Option Explicit


Private Sub LST_PrintNames_Click()
    ClickListShowHide Me.LST_PrintNames.ListIndex
End Sub

Private Sub OPT_PrintShowHoriz_Click()
    ClickOptionsShowHide Me.LST_PrintNames.ListIndex
End Sub

Private Sub OPT_PrintShowVerti_Click()
    ClickOptionsShowHide Me.LST_PrintNames.ListIndex
End Sub

Private Sub OPT_Hide_Click()
    ClickOptionsShowHide Me.LST_PrintNames.ListIndex
End Sub

Private Sub CMD_PrintBack_Click()
    Me.Hide
End Sub

Private Sub CMD_PrintLL_Click()
    Dim sh As Worksheet

    Set sh = ActiveSheet

    On Error Resume Next
    Application.PrintCommunication = False

    With sh.PageSetup
        .LeftMargin = Application.InchesToPoints(0.04)
        .RightMargin = Application.InchesToPoints(0.04)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0.31)
        .FooterMargin = Application.InchesToPoints(0.31)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintTitleRows = "$5:$8"
        .PrintTitleColumns = vbNullString
        .PrintComments = xlPrintNoComments
        .PrintNotes = False
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .FirstPageNumber = xlAutomatic
        .ORDER = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 90
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = sh.ListObjects(1).Range.Address
        .PrintErrors = xlPrintErrorsBlank
    End With
    Application.PrintCommunication = True
    On Error GoTo 0

    Me.Hide
    sh.PrintPreview
End Sub

Private Sub CMD_ColWidth_Click()
    ClickColWidth Me.LST_PrintNames.ListIndex
End Sub

Private Sub CMD_MatchLLShowHide_Click()
    ClickMatchLinelistShowHide
    Me.Hide
End Sub
