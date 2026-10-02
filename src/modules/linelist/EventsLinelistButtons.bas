Attribute VB_Name = "EventsLinelistButtons"
Attribute VB_Description = "Events associated to eventual buttons in the Linelist"
Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events associated to eventual buttons in the Linelist")


Private Const LLSHEET As String = "LinelistTranslation"
Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"
Private Const EXPORTSHEET As String = "Exports"
Private Const PRINTPREFIX As String = "print_"
Private Const CRFPREFIX As String = "crf_"
Private Const TEMPSHEET As String = "temp__"
Private Const SHOWHIDESHEET As String = "show_hide__"

Private showHideObject As IShowHideManager
Private activeShowHideForm As Object
Private tradsform As ITranslationObject   'Translation of forms
Private tradsmess As ITranslationObject   'Translation of messages
Private pass As IPasswords
Private wb As Workbook
Private lltrads As ILLTranslation
Private wkbNames As IHiddenNames

'Initialize translation of forms object
Private Sub InitializeTrads()
    Set wb = ThisWorkbook
    Set lltrads = LLTranslation.Create(wb.Worksheets(LLSHEET))
    Set tradsmess = lltrads.TransObject()
    Set tradsform = lltrads.TransObject(TranslationOfForms)
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEET))
    Set wkbNames = HiddenNames.Create(wb)
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
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
    Application.cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub

'Resolve the ShowHideWorksheetLayer from a sheet tag
Private Function ResolveShowHideLayer(ByVal shType As String) As Byte
    Select Case shType
    Case "HList"
        ResolveShowHideLayer = ShowHideLayerHList
    Case "HList Print"
        ResolveShowHideLayer = ShowHideLayerPrinted
    Case "VList"
        ResolveShowHideLayer = ShowHideLayerVList
    Case "HList CRF"
        ResolveShowHideLayer = ShowHideLayerCRF
    Case Else
        ResolveShowHideLayer = 0
    End Select
End Function

'Return the base sheet name (without print_/crf_ prefix)
Private Function BaseSheetName(ByVal sh As Worksheet) As String
    Dim shType As String

    shType = SheetTag(sh)

    Select Case shType
    Case "HList Print"
        BaseSheetName = Mid$(sh.Name, Len(PRINTPREFIX) + 1)
    Case "HList CRF"
        BaseSheetName = Mid$(sh.Name, Len(CRFPREFIX) + 1)
    Case Else
        BaseSheetName = sh.Name
    End Select
End Function

'Get the sheet type tag (HiddenNames first, cell fallback for legacy sheets).
Private Function SheetTag(ByVal sh As Worksheet) As String
    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)
    SheetTag = shHn.ValueAsString("sheet_type")
End Function

'Get the table name from worksheet-level HiddenNames.
Private Function TableNameOf(ByVal sh As Worksheet) As String
    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)
    TableNameOf = shHn.ValueAsString("table_name")
End Function

'Get the blank row count from worksheet-level HiddenNames.
Private Function RowCountOf(ByVal sh As Worksheet) As Long
    Dim shHn As IHiddenNames
    Set shHn = HiddenNames.Create(sh)
    RowCountOf = shHn.ValueAsLong("row_count")
End Function

'Apply visibility from a ShowHideManager to a worksheet
Private Sub ApplyShowHideVisibility(ByVal mgr As IShowHideManager, _
                                     ByVal sh As Worksheet, _
                                     ByVal shType As String)
    Dim counter As Long
    Dim posIdx As Long

    For counter = 1 To mgr.EntryCount
        posIdx = mgr.PositionIndex(counter)
        If posIdx > 0 Then
            On Error Resume Next
            Select Case shType
            Case "VList"
                sh.Rows(posIdx).Hidden = mgr.IsHidden(counter)
            Case "HList", "HList Print", "HList CRF"
                sh.Columns(posIdx).Hidden = mgr.IsHidden(counter)
            End Select
            On Error GoTo 0
        End If
    Next
End Sub

'Load saved show/hide state from the persistence sheet
Private Sub LoadShowHideState()
    Dim shSH As Worksheet

    If showHideObject Is Nothing Then Exit Sub

    On Error Resume Next
    Set shSH = wb.Worksheets(SHOWHIDESHEET)
    On Error GoTo 0

    If shSH Is Nothing Then Exit Sub
    If shSH.ListObjects.Count = 0 Then Exit Sub

    showHideObject.ImportPlan shSH.ListObjects(1)
End Sub

'Persist the module-level showHideObject state to show_hide__ sheet
Private Sub PersistShowHideState()
    Dim shSH As Worksheet

    If showHideObject Is Nothing Then Exit Sub

    On Error Resume Next
    Set shSH = wb.Worksheets(SHOWHIDESHEET)
    On Error GoTo 0

    If shSH Is Nothing Then Exit Sub
    If shSH.ListObjects.Count = 0 Then Exit Sub

    showHideObject.ExportPlan shSH.ListObjects(1)
End Sub

'Persist a specific ShowHideManager instance to show_hide__ sheet
Private Sub PersistShowHideManager(ByVal mgr As IShowHideManager)
    Dim shSH As Worksheet

    If mgr Is Nothing Then Exit Sub

    On Error Resume Next
    Set shSH = wb.Worksheets(SHOWHIDESHEET)
    On Error GoTo 0

    If shSH Is Nothing Then Exit Sub
    If shSH.ListObjects.Count = 0 Then Exit Sub

    mgr.ExportPlan shSH.ListObjects(1)
End Sub

'Populate a show/hide form's list control from ShowHideManager entries
Private Sub PopulateShowHideList(ByVal frm As Object)
    Dim listCtrl As Object
    Dim counter As Long

    If frm Is Nothing Then Exit Sub
    If showHideObject Is Nothing Then Exit Sub

    On Error Resume Next
    If frm.Name = "F_ShowHideLL" Then
        Set listCtrl = frm.Controls("LST_LLVarNames")
    Else
        Set listCtrl = frm.Controls("LST_PrintNames")
    End If
    On Error GoTo 0

    If listCtrl Is Nothing Then Exit Sub

    listCtrl.Clear
    For counter = 1 To showHideObject.EntryCount
        listCtrl.AddItem showHideObject.HeaderText(counter)
    Next
End Sub

'@Description("Callback for click on show/hide in a linelist worksheet on a button")
'@EntryPoint
Public Sub ClickShowHide()
    Attribute ClickShowHide.VB_Description = "Callback for click on show/hide in a linelist worksheet on a button"

    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim shType As String
    Dim layer As Byte
    Dim frm As Object

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    If (shType <> "HList" And shType <> "HList Print" And shType <> "VList" _
       And shType <> "HList CRF") Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    InitializeTrads

    layer = ResolveShowHideLayer(shType)
    If layer = 0 Then Exit Sub

    Set dict = LLdictionary.Create(ThisWorkbook.Worksheets(DICTSHEET), 1, 1)
    Set showHideObject = ShowHideManager.Create(dict, layer, BaseSheetName(sh))

    'Load saved show/hide state from persistence sheet
    LoadShowHideState

    'Show the appropriate form
    If shType = "HList Print" Or shType = "HList CRF" Then
        Set frm = F_ShowHidePrint
    Else
        Set frm = F_ShowHideLL
    End If

    Set activeShowHideForm = frm
    PopulateShowHideList frm
    frm.Show

    'After form closes, persist state
    PersistShowHideState
    Set activeShowHideForm = Nothing
End Sub

'@Description("Callback for click on the list of showhide")
'@EntryPoint
Public Sub ClickListShowHide(ByVal Index As Long)
    Attribute ClickListShowHide.VB_Description = "Callback for click on the list of showhide"

    Dim entryIdx As Long
    Dim isMand As Boolean

    If showHideObject Is Nothing Then Exit Sub
    If activeShowHideForm Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideObject.EntryCount Then Exit Sub

    isMand = showHideObject.IsMandatory(entryIdx)

    If activeShowHideForm.Name = "F_ShowHideLL" Then
        If showHideObject.IsHidden(entryIdx) Then
            activeShowHideForm.OPT_Hide.Value = True
        Else
            activeShowHideForm.OPT_Show.Value = True
        End If
        activeShowHideForm.OPT_Show.Enabled = Not isMand
        activeShowHideForm.OPT_Hide.Enabled = Not isMand
    Else
        If showHideObject.IsHidden(entryIdx) Then
            activeShowHideForm.OPT_Hide.Value = True
        Else
            activeShowHideForm.OPT_PrintShowHoriz.Value = True
        End If
        activeShowHideForm.OPT_PrintShowHoriz.Enabled = Not isMand
        activeShowHideForm.OPT_PrintShowVerti.Enabled = Not isMand
        activeShowHideForm.OPT_Hide.Enabled = Not isMand
    End If
End Sub

'@Description("Callback for clik on differents show hide options on a button")
'@EntryPoint
Public Sub ClickOptionsShowHide(ByVal Index As Long)
    Attribute ClickOptionsShowHide.VB_Description = "Callback for clik on differents show hide options on a button"

    Dim entryIdx As Long
    Dim shouldHide As Boolean
    Dim sh As Worksheet
    Dim shType As String
    Dim posIdx As Long

    If showHideObject Is Nothing Then Exit Sub
    If activeShowHideForm Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideObject.EntryCount Then Exit Sub
    If showHideObject.IsMandatory(entryIdx) Then Exit Sub

    shouldHide = activeShowHideForm.OPT_Hide.Value
    showHideObject.SetHidden entryIdx, shouldHide

    Set sh = ActiveSheet
    shType = SheetTag(sh)
    posIdx = showHideObject.PositionIndex(entryIdx)

    If posIdx > 0 Then
        pass.UnProtect sh.Name
        On Error Resume Next
        Select Case shType
        Case "VList"
            sh.Rows(posIdx).Hidden = shouldHide
        Case "HList", "HList Print", "HList CRF"
            sh.Columns(posIdx).Hidden = shouldHide
        End Select
        On Error GoTo 0
        pass.Protect sh.Name
    End If
End Sub

'@Description("Callback for click on column width in show/hide")
'@EntryPoint
Public Sub ClickColWidth(ByVal Index As Long)
    Attribute ClickColWidth.VB_Description = "Callback for click on column width in show/hide"

    Dim entryIdx As Long
    Dim posIdx As Long
    Dim sh As Worksheet
    Dim inputValue As String
    Dim colWidth As Long

    If showHideObject Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideObject.EntryCount Then Exit Sub

    posIdx = showHideObject.PositionIndex(entryIdx)
    If posIdx = 0 Then Exit Sub

    InitializeTrads

    Do While True
        inputValue = InputBox(tradsmess.TranslatedValue("MSG_ColWidth"), _
                             tradsmess.TranslatedValue("MSG_Enter"))
        If inputValue = vbNullString Then Exit Sub
        If IsNumeric(inputValue) Then Exit Do
        If MsgBox(tradsmess.TranslatedValue("MSG_EnterNumeric"), _
                  vbOKCancel, vbNullString) = vbCancel Then Exit Sub
    Loop

    colWidth = CLng(inputValue)
    Set sh = ActiveSheet

    pass.UnProtect sh.Name
    On Error Resume Next
    sh.Columns(posIdx).ColumnWidth = colWidth
    On Error GoTo 0
    pass.Protect sh.Name
End Sub


'@Description("Callback for click on the Print Button")
'@EntryPoint
Public Sub ClickOpenPrint()
    Attribute ClickOpenPrint.VB_Description = "Callback for click on the Print Button"

    Dim sh As Worksheet
    Dim printsh As Worksheet
    Dim shType As String

    On Error GoTo ErrOpen

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    InitializeTrads

    If shType <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)
    'UnProtect current workbook
    pass.UnProtect wb
    'Unhide the linelist Print
    printsh.Visible = xlSheetVisible
    printsh.Activate

ErrOpen:
    pass.Protect wb
End Sub


'@Description("Callback for click on the CRF Button")
'@EntryPoint
Public Sub ClickOpenCRF()
    Attribute ClickOpenPrint.VB_Description = "Callback for click on the CRF Button"

    Dim sh As Worksheet
    Dim crfsh As Worksheet
    Dim shType As String

    On Error GoTo ErrOpen

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    InitializeTrads

    If shType <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    Set crfsh = wb.Worksheets(CRFPREFIX & sh.Name)

    'UnProtect current workbook
    pass.UnProtect wb
    'Unhide the linelist Print
    crfsh.Visible = xlSheetVisible
    crfsh.Activate

ErrOpen:
    pass.Protect wb
End Sub

'@Description("Callback for click on close print sheet")
'@EntryPoint
Public Sub ClickClosePrint()
    Attribute ClickClosePrint.VB_Description = "Callback for click on close print/crf sheet"

    Dim sh As Worksheet
    Dim shType As String
    Dim printsh As Worksheet
    Dim crfsh As Worksheet

    On Error GoTo ErrClose
    Set sh = ActiveSheet

    InitializeTrads

    shType = SheetTag(sh)

    If shType <> "HList" And shType <> "HList Print" And shType <> "HList CRF" Then
        WarningOnSheet "MSG_PrintCRFOrDataSheet"
        Exit Sub
    End If

    'Unprotect workbook
    pass.UnProtect wb
    
    If shType = "HList" Then
        Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)
        Set crfsh = wb.Worksheets(CRFPREFIX & sh.Name)
        printsh.Visible = xlSheetVeryHidden
        crfsh.Visible = xlSheetVeryHidden
    ElseIf shType = "HList Print" Then
        Set printsh = sh
        printsh.Visible = xlSheetVeryHidden
    ElseIf shType = "HList CRF" Then
        Set crfsh = sh
        crfsh.Visible = xlSheetVeryHidden
    End If


ErrClose:
    pass.Protect wb
End Sub

'@Description("Rotate all headers in the Print sheet")
'@EntryPoint
Public Sub ClickRotateAll()
    Attribute ClickRotateAll.VB_Description = "Rotate all headers in the Print sheet"

    Dim sh As Worksheet
    Dim Lo As listObject
    Dim hRng As Range
    Dim cRng As Range
    Dim shType As String
    Dim actualOrientation As xlOrientation

    Set sh = ActiveSheet

    shType = SheetTag(sh)

    InitializeTrads

    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    If shType = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)

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
    Dim shType As String
    Dim inputValue As String
    Dim actualRowHeight As Long

    Set sh = ActiveSheet

    shType = SheetTag(sh)

    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    On Error GoTo ErrHand

    InitializeTrads
    BusyApp cursor:=xlNorthwestArrow

    If shType = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)
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
    Dim shType As String

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    If shType <> "HList" And shType <> "HList Print" Then
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
    Dim shType As String
    Dim nbRows As Long

    On Error GoTo errAddRows
    BusyApp cursor:=xlNorthwestArrow
    InitializeTrads
    pass.UnProtect "_active"

    'Unprotect and sending everything
    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    Application.EnableEvents = False

    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)
    nbRows = IIf(shType = "HList", 199, 10)
    csTab.AddRows nbRows:=nbRows
    
    NotBusyApp
    Application.EnableEvents = True
    'Protect only HList

    If shType = "HList" Then pass.Protect "_active"
    Exit Sub

errAddRows:
    On Error Resume Next
    If shType = "HList" Then pass.Protect "_active"
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
    Dim shType As String
    Dim nbBlank As Long

    On Error GoTo errDelRows
    BusyApp cursor:=xlWait
    InitializeTrads
    pass.UnProtect "_active"

    'Unprotect and sending everything
    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    Application.EnableEvents = False

    nbBlank = RowCountOf(sh)
    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)

    csTab.RemoveRows totalCount:=nbBlank

    Application.EnableEvents = True
    NotBusyApp
    If shType = "HList" Then pass.Protect "_active"
    Exit Sub

errDelRows:
    On Error Resume Next
    If shType = "HList" Then pass.Protect "_active"
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

    Const COMMANDHEIGHT As Integer = 25
    Const COMMANDGAPS As Byte = 6

    Dim exportNumber As Integer
    Dim topPosition As Integer
    Dim expObj As ILLExport
    Dim expsh As Worksheet
    Dim totalNumberOfExports As Long
    Dim controlCommand As String
    Dim cmdArray As BetterArray 'List of controls
    Dim btn As MSForms.CommandButton
    Dim btnObj As ExportButton

    'initialize translations
    InitializeTrads

    Set expsh = ThisWorkbook.Worksheets(EXPORTSHEET)
    Set expObj = LLExport.Create(expsh)
    Set cmdArray = New BetterArray
        
    totalNumberOfExports = expObj.NumberOfExports()
    topPosition = COMMANDGAPS

    On Error GoTo errLoadExp

    'Dynamically create and add buttons to the form
        
        For exportNumber = 1 To totalNumberOfExports
            'Add the control if not initialized  
            If expObj.IsActive(exportNumber) Then
                Set btn = F_Export.Controls.Add("Forms.CommandButton.1", "CMDExport" & exportNumber, True)
                btn.Caption = expObj.Value("label button", exportNumber)
                btn.Top = topPosition
                btn.height = COMMANDHEIGHT
                btn.width = 160
                btn.Left = 20
                btn.WordWrap = True
                topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS
                Set btnObj = ExportButton.Create(ThisWorkbook, tradsmess, btn, F_Export.CHK_ExportFiltered)
                cmdArray.Push btnObj
                Set btnObj = Nothing
            End If
        Next

    With F_Export
        'Overall height and width of the form and other parts of the form ------
    
        'Height of checks (use filtered data)
        .CHK_ExportFiltered.Top = topPosition + 30
        .CHK_ExportFiltered.Left = 40
        .CHK_ExportFiltered.width = 160
        topPosition = topPosition + 40 + COMMANDHEIGHT + COMMANDGAPS

        'Height of command for new key
        .CMD_NewKey.Top = topPosition
        .CMD_NewKey.height = COMMANDHEIGHT
        .CMD_NewKey.width = 160
        .CMD_NewKey.Left = 20

        topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS

        'Show Private key command
        .CMD_ShowKey.Top = topPosition
        .CMD_ShowKey.height = COMMANDHEIGHT - 0.1 * COMMANDHEIGHT
        .CMD_ShowKey.width = 160
        .CMD_ShowKey.Left = 20

        topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS

        'Quit command
        .CMD_Back.Top = topPosition
        .CMD_Back.height = COMMANDHEIGHT - 0.1 * COMMANDHEIGHT
        .CMD_Back.width = 160
        .CMD_Back.Left = 20

        topPosition = topPosition + COMMANDHEIGHT + COMMANDGAPS

        .Height = topPosition + 50
        .Width = 210

        'Show the form
        .Show
    End With

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
    Dim shType As String
    Dim sh As Worksheet
    Dim startRow As Long
    Dim tabName As String
    Dim rngName As String

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    If (shType <> "HList") And (shType <> "SPT-Analysis") Then
        WarningOnSheet "MSG_DataOrSpatioSheet"
        Exit Sub
    End If

    InitializeTrads
    
    Select Case shType

    Case "HList"

        tabName = TableNameOf(sh)
        startRow = sh.Range(tabName & "_START").Row + 1
        targetColumn = ActiveCell.Column

        If ActiveCell.Row >= startRow Then

            hfOrGeo = ActiveSheet.Cells(startRow - 5, targetColumn).Value
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
    Dim sheetName As String
    Dim anaSheetsList As BetterArray
    Static previousHit As Long
    Dim timePeriod As Long
    Dim counter As Long

    If (previousHit = 0) Then 
        previousHit = Now()
        timePeriod = 0
    Else
        timePeriod = Now() - previousHit
    End If

    'If the duration is less than 2 minutes, Exit to avoid mutiple recomputations
    If (timePeriod < 60 & timePeriod > 0) Then Exit Sub
    
    InitializeTrads
    On Error GoTo ErrHand

    'Calculate
    BusyApp
    UpdateSpTables

    Set anaSheetsList = New BetterArray
    anaSheetsList.Push wkbNames.ValueAsString("RNG_UASheet"), _
                       wkbNames.ValueAsString("RNG_TSSheet"), _
                       wkbNames.ValueAsString("RNG_SPSheet"), _
                       wkbNames.ValueAsString("RNG_SPTSheet")

    For counter = anaSheetsList.LowerBound To anaSheetsList.UpperBound
        sheetName = anaSheetsList.Item(counter)
        Set sh = wb.Worksheets(sheetName)
        sh.UsedRange.calculate
        sh.Columns("A:E").calculate
    Next

ErrHand:
    NotBusyApp
End Sub


'@Description("Print the current linelist")
'@EntryPoint
Public Sub ClickPrintLL()
    Attribute ClickPrintLL.VB_Description = "Print the current linelist"

    Dim sh As Worksheet
    Dim shType As String

    'Set up the sheet with some print Characteristics
    Set sh = ActiveSheet

    'Test to be sure we are on print or linelist worksheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet
    If shType <> "HList Print" And shType <> "HList" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    'On HListSheet, open the print sheet
    If shType = "HList" Then ClickOpenPrint

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
    Set actsh = wb.Worksheets(wkbNames.ValueAsString("RNG_CustomPivot"))
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
    Dim shType As String
    Dim tabName As String
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
    shType = SheetTag(sh)

    If shType <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    InitializeTrads

    tabName = TableNameOf(sh)
    startRow = sh.Range(tabName & "_START").Row + 1
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

    InitializeTrads
    HandleAnalysisExport wb, tradsmess
End Sub

'@Description("Import new data in the linelist")
'@EntryPoint
Public Sub ClickImportData()
    Attribute clickImportData.VB_Description = "Import new data in the linelist"

    Dim sh As Worksheet
    Dim csTab As ICustomTable
    Dim Lo As ListObject
    Dim nbBlank As Long

    InitializeTrads

    'Resize all HList tables before import (remove blank rows)
    Application.EnableEvents = False
    For Each sh In wb.Worksheets
        If SheetTag(sh) = "HList" Then
            nbBlank = RowCountOf(sh)
            Set Lo = sh.ListObjects(1)
            Set csTab = CustomTable.Create(Lo)
            pass.UnProtect sh.Name
            On Error Resume Next
            If Not (Lo.AutoFilter Is Nothing) Then Lo.AutoFilter.ShowAllData
            csTab.RemoveRows totalCount:=nbBlank
            On Error GoTo 0
            pass.Protect sh.Name
        End If
    Next

    'Import data using LLImporter API (handles file picker, busy state, report)
    HandleImportData wb, tradsmess, False

    'Update all the listAuto in the workbook
    LinelistEventsManager.UpdateAllListAuto

    Application.EnableEvents = True
End Sub

'@Description("Import a new geobase in the linelist")
'@EntryPoint
Public Sub ClickImportGeobase()
    Attribute clickImportGeobase.VB_Description = "Import a new geobase in the linelist"

    InitializeTrads
    HandleImportGeobase wb, tradsmess
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
        sheetInfo = vars.Value(colName:="sheet type", varName:=varName)
        Set sh = ThisWorkbook.Worksheets(varSheetName)

        'On VList, hide the line, on Hlist the column
        If (sheetInfo = "vlist1D") Then
            sh.Cells(varColumnIndex, 1).EntireRow.Hidden = hideStatus
        ElseIf sheetInfo = "hlist2D" Then
            sh.Cells(1, varColumnIndex).EntireColumn.Hidden = hideStatus
        End If

        'Rotate on print sheet
        If (varStatus = "showverti") Then
            tabName = TableNameOf(sh)
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
    Dim mgr As IShowHideManager
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim shType As String
    Dim layer As Byte

    On Error GoTo ErrHand

    BusyApp
    InitializeTrads

    showOptional = (wkbNames.ValueAsString(RNGSHOWALLOPTIONALS) = "yes")

    'Warn user: they will lose current shown/hidden status
    checkConfirm = (MsgBox(tradsmess.TranslatedValue("MSGB_WarningShowHide"), _
                           vbYesNo + vbExclamation, _
                           tradsmess.TranslatedValue("MSGB_Warning")) = vbYes)

    If Not checkConfirm Then GoTo ErrHand

    Set sh = ActiveSheet
    shType = SheetTag(sh)
    layer = ResolveShowHideLayer(shType)
    If layer = 0 Then GoTo ErrHand

    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
    Set mgr = ShowHideManager.Create(dict, layer, BaseSheetName(sh))

    'Toggle all optional entries
    mgr.SetAllOptionalHidden Not showOptional

    'Apply visibility to worksheet
    pass.UnProtect sh.Name
    ApplyShowHideVisibility mgr, sh, shType
    pass.Protect sh.Name

    'Toggle the flag
    If showOptional Then
        wkbNames.SetValue RNGSHOWALLOPTIONALS, "no"
    Else
        wkbNames.SetValue RNGSHOWALLOPTIONALS, "yes"
    End If

    PersistShowHideManager mgr

    'Sync module-level state so ClickShowHide persist is coherent
    Set showHideObject = mgr
    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If

ErrHand:
    NotBusyApp
End Sub



'@Description("Match the show/hide state in the linelist from the print sheet")
'@EntryPoint
Public Sub ClickMatchLinelistShowHide()
    Attribute ClickMatchLinelist.VB_Description = "Match the show/hide state in the linelist from the print sheet"

    Dim checkConfirm As Boolean
    Dim baseMgr As IShowHideManager
    Dim printMgr As IShowHideManager
    Dim printsh As Worksheet
    Dim sh As Worksheet
    Dim dict As ILLdictionary
    Dim sheetName As String
    Dim tabName As String
    Dim counter As Long
    Dim fieldKey As String
    Dim printIdx As Long
    Dim posIdx As Long
    Dim labCellRng As Range

    On Error GoTo ErrHand

    BusyApp
    InitializeTrads

    'Warn user: they will lose current shown/hidden status
    checkConfirm = (MsgBox(tradsmess.TranslatedValue("MSGB_WarningShowHide"), _
                           vbYesNo + vbExclamation, _
                           tradsmess.TranslatedValue("MSGB_Warning")) = vbYes)

    If Not checkConfirm Then GoTo ErrHand

    Set printsh = ActiveSheet
    sheetName = Mid$(printsh.Name, Len(PRINTPREFIX) + 1)
    Set sh = wb.Worksheets(sheetName)
    tabName = TableNameOf(sh)

    Set dict = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
    Set baseMgr = ShowHideManager.Create(dict, ShowHideLayerHList, sheetName)
    Set printMgr = ShowHideManager.Create(dict, ShowHideLayerPrinted, sheetName)

    'Match print visibility to base sheet visibility
    pass.UnProtect printsh.Name
    For counter = 1 To baseMgr.EntryCount
        fieldKey = baseMgr.FieldKey(counter)
        printIdx = printMgr.IndexOf(fieldKey)
        If printIdx > 0 Then
            posIdx = printMgr.PositionIndex(printIdx)
            If posIdx > 0 Then
                If baseMgr.IsHidden(counter) Then
                    printMgr.SetHidden printIdx, True
                    printsh.Columns(posIdx).Hidden = True
                Else
                    printMgr.SetHidden printIdx, False
                    printsh.Columns(posIdx).Hidden = False
                    'Reset orientation to horizontal
                    On Error Resume Next
                    Set labCellRng = printsh.Range( _
                        Replace(tabName, "pr", vbNullString) & "_PRINTSTART")
                    Set labCellRng = printsh.Cells( _
                        labCellRng.Offset(-2).Row, posIdx)
                    labCellRng.Orientation = 0
                    On Error GoTo 0
                End If
            End If
        End If
    Next
    pass.Protect printsh.Name

    PersistShowHideManager baseMgr
    PersistShowHideManager printMgr

    'Sync module-level state so ClickShowHide persist is coherent
    Set showHideObject = printMgr
    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If

ErrHand:
    NotBusyApp
End Sub

'@Description("AutoFit columns/rows in a linelist worksheet")
'@EntryPoint
Public Sub clickAutoFit()
    Attribute clickAutoFit.VB_Description = "AutoFit columns/rows in a linelist worksheet"
    
    Dim sh As Worksheet
    Dim shType As String
    Dim Lo As ListObject
    Dim LoRng As Range
    Dim counter As Long

    On Error GoTo ErrHand

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    If shType <> "HList" Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    BusyApp
    'Table data entry on linelist
    Set Lo = sh.ListObjects(1)

    For counter = 1 To Lo.ListColumns.Count
        Set LoRng = Lo.ListColumns(counter).Range
        
        If (Not LoRng.EntireColumn.Hidden) Then
            'Autofit the column after wrapping the text
            LoRng.WrapText = True
            LoRng.EntireColumn.AutoFit
        End If
    Next
ErrHand:
    On Error Resume Next
    NotBusyApp    
End Sub