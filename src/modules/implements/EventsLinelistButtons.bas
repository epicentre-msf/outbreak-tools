Attribute VB_Name = "EventsLinelistButtons"
Attribute VB_Description = "Events associated to eventual buttons in the Linelist"
Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events associated to eventual buttons in the Linelist")


Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"
Private Const EXPORTSHEET As String = "Exports"
Private Const PRINTPREFIX As String = "print_"
Private Const TEMPSHEET As String = "temp__"
Private Const SHOWHIDESHEET As String = "show_hide__"
Private Const UPDATESHEET As String = "updates__"

Private showHideObject As ILLShowHide
Private tradsform As ITranslation   'Translation of forms
Private tradsmess As ITranslation   'Translation of messages
Private pass As ILLPasswords
Private wb As Workbook
Private lltrads As ILLTranslations

'Initialize translation of forms object
Private Sub InitializeTrads()
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
    MsgBox tradsmess.TranslatedValue(msgCode), vbOKOnly + vbExclamation
End Sub

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub

'@Description("Callback for click on show/hide in a linelist worksheet on a button")
'@EntryPoint
Public Sub ClickShowHide()
    Attribute ClickShowHide.VB_Description = "Callback for click on show/hide in a linelist worksheet on a button"

    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim sheetTag As String
    Dim showOptional As Boolean
    Dim upObj As IUpVal

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
    Set upObj = UpVal.Create(ThisWorkbook.Worksheets(UPDATESHEET))

    'This is the private show hide object, used in future subs.
    Set showHideObject = LLShowHide.Create(tradsmess, dict, sh)
    showOptional = (upObj.Value("RNG_ShowAllOptionals") = "yes")

    'Load elements to the current form
    showHideObject.Load tradsform:=tradsform, showOptional:=showOptional, showForm:=True
End Sub

'@Description("Callback for click on the list of showhide")
'@EntryPoint
Public Sub ClickListShowHide(ByVal Index As Long)
    Attribute ClickListShowHide.VB_Description = "Callback for click on the list of showhide"
    showHideObject.UpdateVisibilityStatus Index
End Sub

'@Description("Callback for clik on differents show hide options on a button")
'@EntryPoint
Public Sub ClickOptionsShowHide(ByVal Index As Long)
    Attribute ClickOptionsShowHide.VB_Description = "Callback for clik on differents show hide options on a button"
    showHideObject.ShowHideLogic Index
End Sub

'@Description("Callback for click on column width in show/hide")
'@EntryPoint
Public Sub ClickColWidth(ByVal Index As Long)
    Attribute ClickColWidth.VB_Description = "Callback for click on column width in show/hide"
    showHideObject.ChangeColWidth Index
End Sub


'@Description("Callback for click on the Print Button")
'@EntryPoint
Public Sub ClickOpenPrint()
    Attribute ClickOpenPrint.VB_Description = "Callback for click on the Print Button"

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
    pass.UnProtectWkb wb
    'Unhide the linelist Print
    printsh.Visible = xlSheetVisible
    printsh.Activate

ErrOpen:
    pass.ProtectWkb wb
End Sub

'@Description("Callback for click on close print sheet")
'@EntryPoint
Public Sub ClickClosePrint()
    Attribute ClickClosePrint.VB_Description = "Callback for click on close print sheet"

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
    pass.UnProtectWkb wb
    printsh.Visible = xlSheetVeryHidden

ErrClose:
    pass.ProtectWkb wb
End Sub

'@Description("Rotate all headers in the Print sheet")
'@EntryPoint
Public Sub ClickRotateAll()
    Attribute ClickRotateAll.VB_Description = "Rotate all headers in the Print sheet"

    Dim sh As Worksheet
    Dim Lo As listObject
    Dim hRng As Range
    Dim cRng As Range
    Dim sheetTag As String
    Dim actualOrientation As xlOrientation

    Set sh = ActiveSheet

    sheetTag = sh.Cells(1, 3).Value

    InitializeTrads

    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    If sheetTag = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)

    On Error GoTo ErrHand

    'Unprotect the sheet if it is protected.
    pass.UnProtect sh.Name
    BusyApp cursor:=xlNorthwestArrow

    Set Lo = sh.ListObjects(1)
    Set hRng = Lo.HeaderRowRange.Offset(-1)
    actualOrientation = IIf(hRng.Orientation = xlUpward, xlHorizontal, xlUpward)
    hRng.Orientation = actualOrientation
    hRng.RowHeight = 100
    
    'AutoFit only non hidden columns
    For Each cRng In hRng
        If Not cRng.EntireColumn.HIDDEN Then cRng.EntireColumn.AutoFit
    Next

ErrHand:
    NotBusyApp
End Sub

'@Description("Change the Row height of cells in the print sheet")
'@EntryPoint
Public Sub ClickRowHeight()
    Attribute ClickRowHeight.VB_Description = "Change the Row height of cells in the print sheet"

    Dim sh As Worksheet
    Dim Lo As listObject
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

    If sheetTag = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)
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
        If (MsgBox(tradsmess.TranslatedValue("MSG_EnterNumeric"), _
             vbOkCancel, vbNullString) = vbCancel) Then Exit Sub
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
    Attribute ClickRemoveFilters.VB_Description = "Click on show all filters"

    Dim sh As Worksheet
    Dim Lo As listObject
    Dim sheetTag As String

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" And sheetTag <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    InitializeTrads
    Set Lo = sh.ListObjects(1)
    On Error GoTo ErrHand

    If Not (Lo.AutoFilter Is Nothing) Then
        BusyApp cursor:=xlNorthwestArrow
        'Unprotect current worksheet
        pass.UnProtect "_active"
        'remove the filters
        Lo.AutoFilter.ShowAllData
        pass.Protect "_active"
    End If
ErrHand:
    pass.Protect "_active"
    NotBusyApp
End Sub

'@Description("Add rows to a data entry table in the Linelist")
'@EntryPoint
Public Sub ClickAddRows()
    Attribute ClickAddRows.VB_Description = "Add rows to a data entry table in the Linelist"

    Dim Lo As listObject
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
    nbRows = IIf(sheetTag = "HList", 199, 10)
    csTab.AddRows nbRows:=nbRows
    
    NotBusyApp
    Application.EnableEvents = True
    'Protect only HList

    If sheetTag = "HList" Then pass.Protect "_active"
    Exit Sub

errAddRows:
    On Error Resume Next
    If sheetTag = "HList" Then pass.Protect "_active"
    NotBusyApp
    Application.EnableEvents = True
    MsgBox tradsmess.TranslatedValue("MSG_ErrAddRows"), _
          vbOKOnly + vbCritical, _
          tradsmess.TranslatedValue("MSG_Error")
    On Error GoTo 0
End Sub

'@Description("Resize the data entry table in the linelist")
'@EntryPoint
Public Sub ClickResize()
    Attribute ClickResize.VB_Description = "Resize the data entry table in the linelist"

    Dim Lo As listObject
    Dim csTab As ICustomTable
    Dim sh As Worksheet
    Dim sheetTag As String
    Dim nbBlank As Long

    On Error GoTo errDelRows
    BusyApp cursor:=xlWait
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
    On Error Resume Next
    If sheetTag = "HList" Then pass.Protect "_active"
    NotBusyApp
    Application.EnableEvents = True
    MsgBox tradsmess.TranslatedValue("MSG_ErrDelRows"), _
          vbOKOnly + vbCritical, _
          tradsmess.TranslatedValue("MSG_Error")
    On Error GoTo 0

End Sub

'@Description("Callback for click on advance configurations")
'@EntryPoint
Public Sub ClickAdvanced()
    Attribute ClickAdvanced.VB_Description = "Callback for click on advance configurations"

    'Import exported data into the linelist
    F_Advanced.Show
End Sub

'@Description("Callback for clik on Export")
'@EntryPoint
Public Sub ClickExport()
    Attribute ClickExport.VB_Description = "Callback for clik on Export"

    Const COMMANDHEIGHT As Integer = 35
    Const COMMANDGAPS As Byte = 10
    Const MAXIMUMNUMBEROFEXPORTS As Integer = 10

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
        For exportNumber = 1 To MAXIMUMNUMBEROFEXPORTS
            If Not exp.IsActive(exportNumber) Then
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
        .width = 200
    End With

    F_Export.Show
    Exit Sub

errLoadExp:
    On Error Resume Next
    MsgBox tradsmess.TranslatedValue("MSG_ErrLoadExport"), _
           vbOKOnly + vbCritical, _
           tradsmess.TranslatedValue("MSG_Error")
    Exit Sub
    On Error GoTo 0
End Sub

'@Description("Callback for clik on open the geobase application")
'@EntryPoint
Public Sub ClickGeoApp()
    Attribute ClickGeoApp.VB_Description = "Callback for clik on open the geobase application"

    Dim targetColumn As Integer
    Dim hfOrGeo As String
    Dim sheetTag As String
    Dim sh As Worksheet
    Dim startRow As Long
    Dim tabName As String
    Dim rngName As String

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If (sheetTag <> "HList") And (sheetTag <> "SPT-Analysis") Then
        WarningOnSheet "MSG_DataOrSpatioSheet"
        Exit Sub
    End If

    InitializeTrads
    
    Select Case sheetTag

    Case "HList"

        tabName = sh.Cells(1, 4).Value
        startRow = sh.Range(TabName & "_" & "START").Row
        targetColumn = ActiveCell.Column

        If ActiveCell.Row >= StartRow Then

            hfOrGeo = ActiveSheet.Cells(StartRow - 5, targetColumn).Value
            Select Case hfOrGeo
            Case "geo1"
                LoadGeo 0
            Case "hf"
                LoadGeo 1
            Case Else
                MsgBox tradsmess.TranslatedValue("MSG_WrongCells")
            End Select
        Else
            MsgBox tradsmess.TranslatedValue("MSG_WrongCells"), _
            vbOKOnly + vbCritical, tradsmess.TranslatedValue("MSG_Error")
        End If

    Case "SPT-Analysis"
        On Error Resume Next
        rngName = ActiveCell.Name.Name
        On Error GoTo 0
        If (InStr(1, rngName, "INPUTSPTGEO_") > 0) Then
            LoadGeo 0
        ElseIf (InStr(1, rngName, "INPUTSPTHF_") > 0) Then
            LoadGeo 1
        End If
    End Select
End Sub

'@Description("Calculate Elements in an analysis worksheet")
'@EntryPoint
Public Sub ClickCalculate()
    Attribute ClickCalculate.VB_Description = "Calculate Elements in an analysis worksheet"

    Dim sh As Worksheet
    Dim sheetTag As String

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "Uni-Bi-Analysis" And sheetTag <> "TS-Analysis" And _ 
       sheetTag <> "SP-Analysis" And sheetTag <> "SPT-Analysis" Then
        WarningOnSheet "MSG_AnaSheet"
        Exit Sub
    End If

    InitializeTrads

    On Error GoTo ErrHand

    'Calculate
    BusyApp

    Select Case sheetTag
    Case "Uni-Bi-Analysis", "TS-Analysis", "SPT-Analysis"
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
    Attribute ClickPrintLL.VB_Description = "Print the current linelist"

    Dim sh As Worksheet
    Dim sheetTag As String

    'Set up the sheet with some print Characteristics
    Set sh = ActiveSheet

    'Test to be sure we are on print or linelist worksheet
    sheetTag = sh.Cells(1, 3).Value

    'Warning if not on print or hlist worksheet
    If sheetTag <> "HList Print" And sheetTag <> "HList" Then
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
        .BottomMargin = Application.InchesToPoints(0.2)
        .HeaderMargin = Application.InchesToPoints(0.31)
        .FooterMargin = Application.InchesToPoints(0.31)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintTitleRows = "$5:$8" 'Those are rows to always keep on title
        .PrintTitleColumns = vbNullString
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
    Attribute ClickExportMigration.VB_Description = "Show the Export for Migration form"

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
        [F_ExportMig].CHK_ExportMigEditableLabel.Value = True
        [F_ExportMig].CHK_ExportMigShowHide.Value = True
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
    Attribute ClickOpenVarLab.VB_Description = "For each table, show the variables and corresponding labels"

    Dim counter As Long 'Counter for the number of tables
    Dim actsh As Worksheet
    Dim tempsh As Worksheet
    Dim dict As ILLdictionary
    Dim varRng As Range
    Dim vars As ILLVariables
    Dim cellRng As Range
    Dim tablename As String
    Dim varLabTab As BetterArray
    Dim varName As String

    InitializeTrads
    On Error GoTo ErrHand

    'Prepare the temporary Sheet
    Set actsh = wb.Worksheets(lltrads.Value("customtable"))
    BusyApp

    Set tempsh = ThisWorkbook.Worksheets(TEMPSHEET)
    tempsh.Cells.Clear

    'Fill in values on pivot table names, and corresponding variables
    Set cellRng = tempsh.Cells(1, 1)

    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
    Set vars = LLVariables.Create(dict)
    
    'Range for variables, as well as final table for the form
    Set varRng = dict.DataRange("variable name")
    Set varLabTab = New BetterArray

    For counter = 1 To varRng.Rows.Count
        varName = varRng.Cells(counter, 1).Value

        'take only variables on sheet of tyme HList, and add them to the table
        If vars.Value(colName:="sheet type", varName:=varName) = "hlist2D" Then

            tablename = vars.Value(colName:="table name", varName:=varName)

            'Pivot table title
            cellRng.Cells(1, 1).Value = actsh.Range("RNG_PivotTitle_" & tablename).Value
            'Varname
            cellRng.Cells(1, 2).Value = varName
            'Corresponding variable label
            cellRng.Cells(1, 3).Value = vars.Value( _
                                        colName:="main label", _
                                        varName:=varName)
            'Move to next line
            Set cellRng = cellRng.Offset(1)
        End If
        'Get the whole table from fill in range
        varLabTab.FromExcelRange tempsh.Cells(1, 1), _
                                 DetectLastRow:=True, DetectLastColumn:=True
    Next

    'Affect the table to the list
    F_ShowVarLabels.LST_CustomTabList.List = varLabTab.Items
    NotBusyApp
    
    'This will open the form with variable name and variable labels for
    [F_ShowVarLabels].Show
ErrHand:
    NotBusyApp
End Sub


'@Description("Sort elements in a current range of a HList worksheet")
'@EntryPoint
Public Sub ClickSortTable()
    Attribute ClickSortTable.VB_Description = "Sort elements in a current range of a HList worksheet"

    Dim sh As Worksheet
    Dim sheetTag As String
    Dim TabName As String
    Dim targetColumn As Long
    Static prevRngName As String
    Static nbTimes As Long
    Dim LoRng As Range
    Dim sortRng As Range
    Dim sortOrder As Long
    Dim StartRow As Long
    Dim headerName As String


    On Error GoTo ErrHand

    Set sh = ActiveSheet
    sheetTag = sh.Cells(1, 3).Value

    If sheetTag <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    InitializeTrads
    
    TabName = sh.Cells(1, 4).Value
    StartRow = sh.Range(TabName & "_" & "START").Row
    targetColumn = ActiveCell.Column
    
    If ActiveCell.Row >= StartRow Then

        headerName = sh.Cells(StartRow - 1, targetColumn).Value

        If (prevRngName <> headerName) Or (nbTimes = 0) Then
            'Ask the user if really want to sort
            prevRngName = headerName
            nbTimes = 0
            If MsgBox( _
                tradsmess.TranslatedValue("MSG_ContinueSort") & " " & headerName, _
                vbYesNo + vbExclamation) = vbNo Then
                Exit Sub
            End If
        End If

        'The sortorder is related to the number of times you clik on the same range
        'For the first time, it is increasing, the second time, decreasing, etc..

        sortOrder = IIf((nbTimes Mod 2) = 0, xlAscending, xlDescending)
        Set LoRng = sh.ListObjects(TabName).Range
        Set sortRng = sh.ListObjects(TabName).ListColumns(headerName).Range
        nbTimes = nbTimes + 1
        
        BusyApp cursor:=xlNorthwestArrow
        'Unprotect the active worksheet, sort the range, and protect back.
        'I have to keep the protect/unprotect step as far as possible for
        'performance issues
        pass.UnProtect "_active"
        On Error Resume Next
        LoRng.Sort key1:=sortRng, Order1:=sortOrder, Header:=xlYes
        On Error GoTo 0
        pass.Protect "_active"
    End If

ErrHand:
    NotBusyApp
End Sub


'@Description("Export all Analysis worksheets to a workbook")
'@EntryPoint
Public Sub ClickExportAnalysis()
    Attribute ClickExportAnalysis.VB_Description = "Export all Analysis worksheets to a workbook"
    
    Dim expOut As IOutputSpecs

    'Add Error management
    On Error GoTo ErrHand
    InitializeTrads

    BusyApp
    Set expOut = OutputSpecs.Create(wb, ExportAna)
    expOut.Save tradsmess
    NotBusyApp
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox tradsmess.TranslatedValue("MSG_ErrHandExport"), _
            vbOKOnly + vbCritical, _
            tradsmess.TranslatedValue("MSG_Error")
    NotBusyApp
    On Error GoTo 0
End Sub

'@Description("Import new data in the linelist")
'@EntryPoint
Public Sub ClickImportData()
    Attribute clickImportData.VB_Description = "Import new data in the linelist"
    
    Dim impObj As IImpSpecs
    Dim sh As Worksheet
    Dim csTab As ICustomTable
    Dim Lo As ListObject
    Dim nbBlank As Long

    InitializeTrads
    BusyApp

    Set impObj = ImpSpecs.Create([F_ImportRep], [F_Advanced], wb)
    
    'Resize all the table on all HList worksheets
    Application.EnableEvents = False
    For Each sh in wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            nbBlank = sh.Cells(1, 6).Value
            Set Lo = sh.ListObjects(1)
            Set csTab = CustomTable.Create(Lo)
            pass.UnProtect sh.Name
            On Error Resume Next
            csTab.RemoveRows totalCount:=nbBlank
            On Error GoTo 0
            pass.Protect sh.Name
        End If
    Next
    
    On Error GoTo ErrManage

    impObj.ImportMigration

    'Update all the listAuto in the workbook
    UpdateAllListAuto wb

ErrManage:
    NotBusyApp
    Application.EnableEvents = True
End Sub

'@Description("Import a new geobase in the linelist")
'@EntryPoint
Public Sub ClickImportGeobase()
    Attribute clickImportGeobase.VB_Description = "Import a new geobase in the linelist"
    
    Dim impObj As IImpSpecs
    Dim currwb As Workbook

    Set currwb = ThisWorkbook
    Set impObj = ImpSpecs.Create([F_ImportRep], [F_Advanced], currwb)
    impObj.ImportGeobase
End Sub

Private Sub RestaureHiddenStatus(ByVal cellRng As Range, Optional ByVal scope As Byte = 1)

    'Scope can take 3 values:
    '1- dictionary
    '2- show/hide, on a linelist
    '3- show/hide, on a linelist Print

    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim sh As Worksheet
    Dim varName As String
    Dim tabName As String
    Dim varSheetName As String
    Dim hideStatus As Boolean
    Dim varStatus As String
    Dim varColumnIndex As Long
    Dim sheetInfo As String
    Dim prevSheetName As String
    Dim labCellRng As Range

    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
    Set vars = LLVariables.Create(dict)

    'Hide all hidden variables

    Do While (Not IsEmpty(cellRng))
        
        varName = cellRng.Value
        varStatus = vbNullString

        If (scope = 1) Then 
            varStatus = vars.Value(colName:="status", varName:=varName)
        Else
            'on the show/hide worksheet the variable status is related
            'to the language
            Select Case cellRng.Offset( ,1).Value

            Case tradsmess.TranslatedValue("MSG_Hidden")
                varStatus = "hidden"
            Case tradsmess.TranslatedValue("MSG_Mandatory")
                varStatus = "mandatory"
            Case tradsmess.TranslatedValue("MSG_Shown")
                varStatus = "shown"
            Case tradsmess.TranslatedValue("MSG_ShowHoriz")
                varStatus = "showhoriz"
            Case tradsmess.TranslatedValue("MSG_ShowVerti")
                varStatus = "showverti"
            End Select
        End If

        varSheetName = vars.Value(colName:="sheet name", varName:=varName)
        
        'On Print sheet, add the print tag to the worksheet name
        If (scope = 3) Then varSheetName = "print_" & varSheetName
        
        varColumnIndex = 0
        On Error Resume Next
            varColumnIndex = CLng(vars.Value(colName:="column index", varName:=varName))
        On Error GoTo 0

        'If unable to identify the column, skip    
        If varColumnIndex = 0 Then GoTo ContinueLoop
        'On mandatory variables in the show/hide worksheet, skip
        If ((scope <> 1) And (varStatus = "mandatory")) Then GoTo ContinueLoop

        'Only hidden variables are hidden, otherwise show them (mandatory or shown)
        hideStatus = IIF(varStatus = "hidden", True, False)

        'Protect or unprotect the sheet of the variable (There is no need to protect/unproctect on print sheet)
        If (prevSheetName <> varSheetName) And (scope <> 3) Then 
            If (prevSheetName <> vbNullString) Then pass.Protect prevSheetName
            Set sh = wb.Worksheets(varSheetName)
            prevSheetName = varSheetName
            pass.UnProtect varSheetName
        End If
        
        sheetInfo = vars.Value(colName:="sheet type", varName:=varName)
        Set sh = ThisWorkbook.Worksheets(varSheetName)

        'On VList, hide the line, on Hlist the column
        If (sheetInfo = "vlist1D") And (scope = 1) Then
            sh.Cells(varColumnIndex, 1).EntireRow.Hidden = hideStatus
        ElseIf sheetInfo = "hlist2D" Then
            sh.Cells(1, varColumnIndex).EntireColumn.Hidden = hideStatus
        End If

        'Rotate on print sheet
        If (varStatus = "showverti") Then
            tabName = sh.Cells(1, 4).Value
            On Error Resume Next
                Set labCellRng = sh.Range(Replace(tabName, "pr", vbNullString) & "_" & "PRINTSTART")
                Set labCellRng = sh.Cells(labCellRng.Offset(-2).Row, varColumnIndex) 
                labCellRng.EntireRow.RowHeight = 100
                labCellRng.Orientation = 90
            On Error GoTo 0
        End If

    ContinueLoop:
        Set cellRng = cellRng.Offset(1)
    Loop

    pass.Protect varSheetName
End Sub

'@Description("Reset hidden columns in the linelist")
'@EntryPoint
Public Sub ClickResetColumns()
    Attribute clickResetColumns.VB_Description = "Reset hidden columns in the linelist"
    
   Dim Lo As ListObject
   Dim cellRng As Range
   Dim tabName As String
   Dim scope As Byte

    BusyApp

    InitializeTrads

    On Error GoTo ErrHand

    'Return the state of variables in the dictionary
    Set cellRng = wb.Worksheets(DICTSHEET).Cells(2, 1)
    RestaureHiddenStatus cellRng, scope:=1

    'Return the state of the variables in show/hide worksheet
    For Each Lo in wb.Worksheets(SHOWHIDESHEET).ListObjects
        'cellRange
        tabName = Replace(Lo.Name, "ShowHideTable_", vbNullString)
        scope = IIF(InStr(1, tabName, "pr") = 1, 3, 2)
        Set cellRng = Lo.Range.Cells(2, 2)
        RestaureHiddenStatus cellRng, scope:=scope
    Next

ErrHand:
    NotBusyApp
End Sub

'@Description("Hide/Unhide Optional variables in the linelist")
'@EntryPoint
Public Sub ClickShowHideMinimal()
    Attribute ClickShowHideMinimal.VB_Description = "Hide/Unhide Optional variables in the linelist"
    
    Const RNGSHOWALLOPTIONALS As String = "RNG_ShowAllOptionals"

    Dim showOptional As Boolean
    Dim checkConfirm As Boolean
    Dim showHideObject As ILLShowHide
    Dim wb As Workbook
    Dim shCsTab As ICustomTable
    Dim varRng As Range
    Dim counter As Long
    Dim varName As String
    Dim varStatus As String
    Dim colIndex As Long
    Dim hiddenShowTag As String
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim vars As ILLVariables
    Dim upObj As IUpVal

    On Error GoTo ErrHand

    BusyApp
    Set wb = ThisWorkbook
    InitializeTrads

    Set upObj = UpVal.Create(wb.Worksheets(UPDATESHEET))
    showOptional = (upObj.Value(RNGSHOWALLOPTIONALS) = "yes")
    
    If showOptional Then
        hiddenShowTag = tradsmess.TranslatedValue("MSG_Shown")
    Else
        hiddenShowTag = tradsmess.TranslatedValue("MSG_Hidden")
    End If

    'Issue a warning to the user: He/She will loose the status of the shown/hidden columns
    checkConfirm = (MsgBox(tradsmess.TranslatedValue("MSGB_WarningShowHide"), _ 
                           vbYesNo + vbExclamation, _ 
                           tradsmess.TranslatedValue("MSGB_Warning")) = vbYes)

    If Not checkConfirm Then GoTo ErrHand

    'Custom table of the show/hide
    Set sh = ActiveSheet
    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
    Set vars = LLVariables.Create(dict)
    Set showHideObject = LLShowHide.Create(tradsmess, dict, sh)      
    Set shCsTab = showHideObject.ShowHideTable()
    Set varRng = shCsTab.DataRange("variable name", strictSearch:=True)

    For counter = 1 To varRng.Rows.Count
        varName = varRng.Cells(counter, 1).Value
        varStatus = vbNullString
        varStatus = vars.Value(varName:=varName, colName:="status")
        
        If (varStatus <> "mandatory") And (varStatus <> vbNullString) And (varStatus <> "hidden") Then
            colIndex = 0
            On Error Resume Next
            colIndex = CLng(vars.Value(varName:=varName, colName:="column index"))

            'Hide the column index of the variable
            If colIndex <> 0 Then
                If showOptional Then
                    sh.Columns(colIndex).ColumnWidth = 22
                Else
                    sh.Columns(colIndex).ColumnWidth = 0
                End If
                'Change the status in the show/hide worksheet
                shCsTab.SetValue colName:="status", keyName:=varName, newValue:=hiddenShowTag
            End If
            On Error GoTo 0
        End If
    Next

    'Update the upObj. If the user wants to show all Optional variables
    'set the show all optionals to no (because previous value was yes)
    'If the user wants to hide all Optional variables, est show optional to "yes"
    '(because previous value was no)
    If showOptional Then
        upobj.SetValue RNGSHOWALLOPTIONALS, "no"
    Else
        upobj.SetValue RNGSHOWALLOPTIONALS, "yes"
    End If

    showHideObject.Load tradsform:=tradsform, showForm:=False, _ 
                        showOptional:=(Not showOptional)

ErrHand:
    NotBusyApp
End Sub



'@Description("Match the show/hide state in the linelist from the print sheet")
'@EntryPoint
Public Sub ClickMatchLinelistShowHide()
    Attribute ClickMatchLinelist.VB_Description = "Match the show/hide state in the linelist from the print sheet"
    
    Dim checkConfirm As Boolean
    Dim showHideObject As ILLShowHide
    Dim showHidePrintObject As ILLShowHide
    Dim wb As Workbook
    Dim shCsTab As ICustomTable
    Dim shPrintCsTab As ICustomTable
    Dim varRng As Range
    Dim counter As Long
    Dim varName As String
    Dim varStatus As String
    Dim colIndex As Long
    Dim hiddenTag As String
    Dim printsh As Worksheet
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim sheetName As String

    On Error GoTo ErrHand

    BusyApp
    Set wb = ThisWorkbook
    InitializeTrads

    'Issue a warning to the user: He/She will loose the status of the shown/hidden columns
    checkConfirm = (MsgBox(tradsmess.TranslatedValue("MSGB_WarningShowHide"), _ 
                           vbYesNo + vbExclamation, _ 
                           tradsmess.TranslatedValue("MSGB_Warning")) = vbYes)

    If Not checkConfirm Then GoTo ErrHand

    hiddenTag = tradsmess.TranslatedValue("MSG_Hidden")

    'Custom table of the show/hide
    Set printsh = ActiveSheet

    '6 is the length of the print_ tag at the begining of the print worksheet name
    sheetName = Right(printsh.Name, (Len(printsh.Name) - 6))
    
    Set sh = wb.Worksheets(sheetName)
    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
    Set showHideObject = LLShowHide.Create(tradsmess, dict, sh)
    Set showHidePrintObject = LLShowHide.Create(tradsmess, dict, printsh)      
    Set shCsTab = showHideObject.ShowHideTable()
    Set shPrintCsTab = showHidePrintObject.ShowHideTable()
    Set varRng = shCsTab.DataRange("variable name", strictSearch:=True)

    For counter = 1 To varRng.Rows.Count
        
        varName = varRng.Cells(counter, 1).Value
        varStatus = vbNullString
        varStatus = shCsTab.Value(keyName:=varName, colName:="status")

        If varStatus = hiddenTag Then
            colIndex = 0
            On Error Resume Next
            colIndex = CLng(shCsTab.Value(keyName:=varName, colName:="column index"))

            'Hide the column index of the variable
            If colIndex <> 0 Then
                printsh.Columns(colIndex).ColumnWidth = 0
                'Change the status in the show/hide worksheet, on the print table
                shPrintCsTab.SetValue colName:="status", keyName:=varName, newValue:=hiddenTag
            End If
            On Error GoTo 0
        End If
    Next

    showHidePrintObject.Load tradsform:=tradsform, showForm:=False
ErrHand:
    NotBusyApp
End Sub