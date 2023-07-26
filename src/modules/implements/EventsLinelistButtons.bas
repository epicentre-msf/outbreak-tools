Attribute VB_Name = "EventsLinelistButtons"
Attribute VB_Description = "Events associated to eventual buttons in the Linelist"
Option Explicit
Option Private Module
Public iGeoType As Byte

'@Folder("Linelist Events")
'@ModuleDescription("Events associated to eventual buttons in the Linelist")


Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"
Private Const EXPORTSHEET As String = "Exports"
Private Const PRINTPREFIX As String = "print_"

Private showHideObject As ILLShowHide
Private tradsform As ITranslation   'Translation of forms
Private tradsmess As ITranslation   'Translation of messages
Private pass As ILLPasswords
Private wb As Workbook

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltrads As ILLTranslations
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet


    Set wb = ThisWorkbook
    Set lltranssh = ThisWorkbook.Worksheets(LLSHEET)
    Set dicttranssh = ThisWorkbook.Worksheets(TRADSHEET)
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradsmess = lltrads.TransObject()
    Set tradsform = lltrads.TransObject(TranslationOfForms)
    Set pass = LLPasswords.Create(wb.Worksheets(PASSSHEET))
End Sub

Private Sub WarningOnSheet(ByVal msgCode As String)
    InitializeTrads
    MsgBox tradsmess.TranslatedValue(msgCode), vbOkOnly + vbExclamation
End Sub

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Cursor = xlDefault
End Sub

'@Description("Callback for click on show/hide in a linelist worksheet on a button")
'@EntryPoint
Public Sub ClickShowHide()
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim sheetTag As String

    Set sh = ActiveSheet
    'Test the sheet type to be sure it is a HList or a HList Print,
    'and exit if not
    sheetTag = sh.Cells(1, 3).Value
    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    'initialize the translations of forms and messages
    InitializeTrads

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets(DICTSHEET), 1, 1)

    'This is the private show hide object, used in future subs.
    Set showHideObject = LLShowHide.Create(tradsmess, dict, sh)

    'Load elements to the current form
    showHideObject.Load tradsform
End Sub

'@Description("Callback for click on the list of showhide")
'@EntryPoint
Public Sub ClickListShowHide(ByVal index As Long)
    showHideObject.UpdateVisibilityStatus index
End Sub

'@Description("Callback for clik on differents show hide options on a button")
'@EntryPoint
Public Sub ClickOptionsShowHide(ByVal index As Long)
    showHideObject.ShowHideLogic index
End Sub

'@Description("Callback for click on column width in show/hide")
'@EntryPoint
Public Sub ClickColWidth(ByVal index As Long)
    showHideObject.ChangeColWidth index
End Sub


'@Description("Callback for click on the Print Button")
'@EntryPoint
Public Sub ClickOpenPrint()

    Dim sh As Worksheet
    Dim printsh As Worksheet
    Dim sheetTag As String

    On Error GoTo ErrOpen

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    InitializeTrads

    If sheetTag <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)
    'UnProtect current workbook
    pass.UnprotectWkb wb
    'Unhide the linelist Print
    printsh.Visible = xlSheetVisible
    printsh.Activate

ErrOpen:
    pass.ProtectWkb wb
End Sub

'@Description("Callback for click on close print sheet")
'@EntryPoint
Public Sub ClickClosePrint()

    Dim sh As Worksheet
    Dim sheetTag As String
    Dim printsh As Worksheet

    On Error GoTo ErrClose
    Set sh = ActiveSheet

    InitializeTrads

    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    If sheetTag = "HList" Then
        Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)
    Else
        Set printsh = sh
    End If
    'Unprotect workbook
    pass.UnprotectWkb wb
    printsh.Visible = xlSheetVeryHidden

ErrClose:
    pass.ProtectWkb wb
End Sub

'@Description("Rotate all headers in the Print sheet")
'@EntryPoint
Public Sub ClickRotateAll()

    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim hRng As Range
    Dim cRng As Range
    Dim sheetTag As String
    Dim actualOrientation As xlOrientation

    Set sh = ActiveSheet

    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    If sheetTag = "HList" Then  Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)

    On Error GoTo ErrHand

    InitializeTrads

    'Unprotect the sheet if it is protected.
    pass.UnProtect sh.Name
    BusyApp cursor:=xlNorthwestArrow

    Set Lo = sh.ListObjects(1)
    Set hRng = Lo.HeaderRowRange.Offset(-1)
    actualOrientation = IIf(hRng.Orientation = xlUpward, xlHorizontal, xlUpward)
    hRng.Orientation = actualOrientation
    hRng.RowHeight = 100
    
    'AutoFit only non hidden columns
    For Each cRng in hRng
        If Not cRng.EntireColumn.Hidden Then cRng.EntireColumn.AutoFit
    Next

ErrHand:
    NotBusyApp
End Sub

'@Description("Change the Row height of cells in the print sheet")
'@EntryPoint
Public Sub ClickRowHeight()

    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim sheetTag As String
    Dim inputValue As String
    Dim actualRowHeight As Long

    Set sh = ActiveSheet

    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    On Error GoTo ErrHand

    InitializeTrads
    BusyApp cursor:=xlNorthwestArrow

    If sheetTag = "HList" Then  Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)
    pass.UnProtect sh.Name

    Set Lo = sh.ListObjects(1)
    If (Lo.DataBodyRange Is Nothing) Then
        Set LoRng = Lo.HeaderRowRange.Offset(1)
    Else
        Set LoRng = Lo.DataBodyRange
    End If

    'Ask for rowheight
    Do While (True)
        inputValue = InputBox(tradsmess.TranslatedValue("MSG_RowHeight"), _
                             tradsmess.TranslatedValue("MSG_Enter"))
        If inputValue = vbNullString Then Exit Sub
        If IsNumeric(inputValue) Then Exit Do
        If (MsgBox (tradsmess.TranslatedValue("MSG_EnterNumeric"), _
             vbOkCancel, "") = vbCancel) Then Exit Sub
    Loop

    On Error Resume Next
        actualRowHeight = CLng(inputValue)
        LoRng.EntireRow.RowHeight = actualRowHeight
    On Error GoTo 0

ErrHand:
    NotBusyApp
End Sub


'@Description("Click on show all filters")
'@EntryPoint
Public Sub ClickRemoveFilters()
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim sheetTag As String

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    InitializeTrads
    Set Lo = sh.ListObjects(1)
    On Error GoTo errHand

    If Not (Lo.AutoFilter Is Nothing) Then
        BusyApp cursor:=xlNorthwestArrow
        'Unprotect current worksheet
        pass.UnProtect "_active"
        'remove the filters
        Lo.AutoFilter.ShowAllData
        pass.Protect "_active"
    End If
ErrHand:
    NotBusyApp
End Sub

'@Description("Add rows to a data entry table in the Linelist")
'@EntryPoint
Public Sub ClickAddRows()

    Dim Lo As ListObject
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Dim sheetTag As String
    Dim nbRows As Long

    On Error GoTo errAddRows
    BusyApp cursor:=xlNorthwestArrow
    InitializeTrads
    pass.UnProtect "_active"

    'Unprotect and sending everything
    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    'Warning if not on print or hlist worksheet
    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    Application.EnableEvents = False

    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)
    nbRows = IIf(sheetTag = "HList", 100, 10)
    csTab.AddRows nbRows:=nbRows
    
    NotBusyApp
    Application.EnableEvents = True
    'Protect only HList

    If sheetTag = "HList" Then pass.Protect "_active"
    Exit Sub

errAddRows:
    NotBusyApp
    Application.EnableEvents = True
    MsgBox tradsmess.TranslatedValue("MSG_ErrAddRows"), _
          vbOKOnly + vbCritical, _
          tradsmess.TranslatedValue("MSG_Error")
    Exit Sub
End Sub

'@Description("Resize the data entry table in the linelist")
'@EntryPoint
Public Sub ClickResize()
    Dim Lo As ListObject
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Dim sheetTag As String
    Dim nbBlank As Long

    On Error GoTo errDelRows
    BusyApp cursor:=xlNorthwestArrow
    InitializeTrads
    pass.UnProtect "_active"

    'Unprotect and sending everything
    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    'Warning if not on print or hlist worksheet
    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    Application.EnableEvents = False

    nbBlank = sh.Cells(1, 6).Value
    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)

    csTab.RemoveRows totalCount:=nbBlank

    Application.EnableEvents = True
    NotBusyApp
    If sheetTag = "HList" Then pass.Protect "_active"
    Exit Sub

errDelRows:
    NotBusyApp
    Application.EnableEvents = True
    MsgBox tradsmess.TranslatedValue("MSG_ErrDelRows"), _
          vbOKOnly + vbCritical, _
          tradsmess.TranslatedValue("MSG_Error")
    Exit Sub

End Sub

'@Description("Callback for click on advance configurations")
'@EntryPoint
Public Sub ClickAdvanced()
    'Import exported data into the linelist
    F_Advanced.Show
End Sub

'@Description("Callback for clik on Export")
'@EntryPoint
Public Sub ClickExport()

    Const COMMANDHEIGHT As Integer = 50
    Const COMMANDGAPS As Byte = 10

    Dim exportNumber As Integer
    Dim topPosition As Integer
    Dim exp As ILLExport
    Dim expsh As Worksheet

    Set expsh = ThisWorkbook.Worksheets(EXPORTSHEET)
    Set exp = LLExport.Create(expsh)

    'initialize translations
    InitializeTrads
    topPosition = COMMANDGAPS

    On Error GoTo errLoadExp

    With F_Export
        For exportNumber = 1 To 5
            If  Not exp.IsActive(exportNumber) Then
                .Controls("CMD_Export" & exportNumber).Visible = False
            Else
                .Controls("CMD_Export" & exportNumber).Visible = True
                .Controls("CMD_Export" & exportNumber).Caption = exp.Value("label button", exportNumber)
                .Controls("CMD_Export" & exportNumber).Top = topPosition
                .Controls("CMD_Export" & exportNumber).height = COMMANDHEIGHT
                .Controls("CMD_Export" & exportNumber).width = 160
                .Controls("CMD_Export" & exportNumber).Left = 20
                .Controls("CMD_Export" & exportNumber).WordWrap = True
                topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS
            End If
        Next

        'Height of checks (use filtered data)
        .CHK_ExportFiltered.Top = topPosition + 30
        .CHK_ExportFiltered.Left = 30
        .CHK_ExportFiltered.width = 160
        topPosition = topPosition + 40 + COMMANDHEIGHT + COMMANDGAPS

        'Height of command for new key
        .CMD_NewKey.Top = topPosition
        .CMD_NewKey.height = COMMANDHEIGHT - 10
        .CMD_NewKey.width = 160
        .CMD_NewKey.Left = 20

        topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS

        'Show Private key command
        .CMD_ShowKey.Top = topPosition
        .CMD_ShowKey.height = COMMANDHEIGHT - 10
        .CMD_ShowKey.width = 160
        .CMD_ShowKey.Left = 20

        topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS

        'Quit command
        .CMD_Back.Top = topPosition
        .CMD_Back.height = COMMANDHEIGHT - 10
        .CMD_Back.width = 160
        .CMD_Back.Left = 20

        topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS

        'Overall height and width of the form

        .height = topPosition + 50
        .width = 210
    End With

    F_Export.Show
    Exit Sub

errLoadExp:
    MsgBox tradsmess.TranslatedValue("MSG_ErrLoadExport"), _
           vbOKOnly + vbCritical, _
           tradsmess.TranslatedValue("MSG_Error")
    Exit Sub
End Sub

'@Description("Callback for clik on open the geobase application")
'@EntryPoint
Public Sub ClickGeoApp()

    Dim targetColumn As Integer
    Dim geoOrhf As String
    Dim sheetTag As String
    Dim sh As Worksheet
    Dim startRow As Long
    Dim tabName As String

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    InitializeTrads
    
    tabName = sh.Cells(1, 4).Value
    startRow = sh.Range(tabName & "_" & "START").Row
    targetColumn = ActiveCell.Column
    

    If ActiveCell.Row >= startRow Then

        geoOrhf = ActiveSheet.Cells(startRow - 5, targetColumn).Value
        Select Case geoOrhf

        Case "geo1"
            iGeoType = 0
            LoadGeo 0

        Case "hf"
            iGeoType = 1
            LoadGeo 1

        Case Else
            MsgBox tradsmess.TranslatedValue("MSG_WrongCells")
        End Select
    Else
        MsgBox tradsmess.TranslatedValue("MSG_WrongCells"),  _ 
        vbOKOnly + vbCritical, tradsmess.TranslatedValue("MSG_Error")
    End If
End Sub

'@Description("Calculate Elements in an analysis worksheet")
'@EntryPoint

Public Sub ClickCalculate()

    Dim sh As Worksheet
    Dim sheetTag As String

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "Uni-Bi-Analysis" And sheetTag <> "TS-Analysis" And sheetTag <> "SP-Analysis" Then
        WarningOnSheet "MSG_AnaSheet"
        Exit Sub
    End If

    InitializeTrads

    On Error GoTo ErrHand

    'Calculate
    BusyApp

    Select Case sheetTag
    Case "Uni-Bi-Analysis"
        UpdateFilterTables
    Case "TS-Analysis"
        UpdateFilterTables
    Case "SP-Analysis"
        UpdateSpTables
    End Select
ErrHand:
    NotBusyApp
End Sub


'@Description("Print the current linelist")
'@EntryPoint
Public Sub ClickPrintLL()
    Dim sh As Worksheet
    Dim sheetTag As String

    'Set up the sheet with some print Characteristics
    Set sh = ActiveSheet

    'Test to be sure we are on print or linelist worksheet
    sheetTag = sh.Cells(1, 3).Value

    'Warning if not on print or hlist worksheet
    If  sheetTag <> "HList Print" And sheetTag <> "HList" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    'On HListSheet, open the print sheet
    If sheetTag = "HList" Then ClickOpenPrint

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
        .Zoom = 100
        .FitToPagesWide = 1
        .FitToPagesTall = False
        'Print Errors to blanks
        .PrintArea = sh.ListObjects(1).Range.Address
        .PrintErrors = xlPrintErrorsBlank
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
    
    sh.PrintPreview
End Sub

'@Description("Show the Export for Migration form")
'@EntryPoint
Public Sub ClickExportMigration()

    'This static variable will keep the selection of
    'the user after the first click. The variable
    'will remain active as long As the workbook is open
    Static AfterFirstClicMig As Boolean

    If AfterFirstClicMig Then
        [F_ExportMig].Show
    Else
        'For the first click Thick Migration and Geo and put historic to false
        'For subsequent clicks, just show what have been ticked
        [F_ExportMig].CHK_ExportMigData.Value = True
        [F_ExportMig].CHK_ExportMigGeo.Value = True
        [F_ExportMig].CHK_ExportMigGeoHistoric.Value = True
        [F_ExportMig].Show
        AfterFirstClicMig = True
    End If
End Sub

'@Description("For each table, show the variables and corresponding labels")
'@EntryPoint
Public Sub ClickOpenVarLab()

    'This will open the form with variable name and variable labels for
    
End Sub
