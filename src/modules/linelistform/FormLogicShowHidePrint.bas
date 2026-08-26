Attribute VB_Name = "FormLogicShowHidePrint"
Attribute VB_Description = "Form code-behind for F_ShowHidePrint"

'@Folder("Linelist Forms")
'@IgnoreModule UnrecognizedAnnotation, UnassignedVariableUsage, UndeclaredVariable
'@ModuleDescription("Form code-behind for F_ShowHidePrint")
'@depends LinelistEventsManager, EventLinelist, LLTranslation, TranslationObject

' This module is the complete code-behind of the F_ShowHidePrint form and is
' copied into the form at deployment. The form is the printed sheet's show/hide:
' beside showing and hiding a variable it says which way the header of a shown
' variable reads, which is the one thing the data entry form has no need of.
'
' The controls the form carries:
'
'   LST_PrintNames        three columns, the label, the variable and its status
'   OPT_PrintShowHoriz    show the picked variable, header across
'   OPT_PrintShowVerti    show the picked variable, header down
'   OPT_Hide              hide the picked variable
'   LBL_ID                header of the variable column
'   LBL_PrintStatus       header of the status column
'   CMD_PrintBack         leave the form
'   CMD_PrintLL           set the page up and open the print preview
'   CMD_ColWidth          change the width of the picked column
'   CMD_MatchLLShowHide   take the choices of the data entry sheet
'
' Every one of them has a row in T_TradLLForms under its own name, so
' UserForm_Initialize translates the whole form in one call.

Option Explicit


Private tradform As TranslationObject


'The translation helper is the one EventLinelist holds.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist
    Dim lltrads As LLTranslation

    If Not tradform Is Nothing Then Exit Sub

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If Not linelistEvents Is Nothing Then Set lltrads = linelistEvents.Translation()

    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "FormLogicShowHidePrint", _
                  "This linelist carries no usable translation sheet"

    Set tradform = lltrads.TransObject(TranslationOfForms)
End Sub

Private Sub UserForm_Initialize()
    InitializeTrads

    Me.Caption = tradform.TranslatedValue(Me.Name)
    tradform.TranslateForm Me
End Sub

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
