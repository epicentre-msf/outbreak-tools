Attribute VB_Name = "EventsLinelistButtons"
Attribute VB_Description = "Events associated to eventual buttons in the Linelist"
Option Explicit
Option Private Module

'@Folder("Linelist Events")
'@ModuleDescription("Events associated to eventual buttons in the Linelist")


Private Const DICTSHEET As String = "Dictionary"
Private Const PASSSHEET As String = "__pass"
Private Const EXPORTSHEET As String = "Exports"
Private Const PRINTPREFIX As String = "print_"
Private Const CRFPREFIX As String = "crf_"

'The pair the show/hide form is open on: which variables the sheet offers, and
'the sheet itself. Both are rebuilt each time the form opens.
Private showHideEntries As ShowHide
Private showHideLayout As ShowHideLayout
Private activeShowHideForm As Object
Private tradsform As TranslationObject   'Translation of forms
Private tradsmess As TranslationObject   'Translation of messages
Private pass As Passwords
Private wb As Workbook
Private lltrads As LLTranslation
Private wkbNames As HiddenNames

'The event service of the running linelist
'
'It holds the translation helper, the workbook hidden names, the password
'manager and one hidden name store per worksheet. This module used to build its
'own of each on every button press: three walks of a Names collection before the
'button had done anything.
'
'The typed local is what call-signature-scan.R reads to check the member. A
'chained call is invisible to it, and this module carries no registry row, so a
'chain here would be checked by nothing at all.
Private Function LinelistService() As EventLinelist
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    Set LinelistService = linelistEvents
End Function

'Initialize translation of forms object
'
'The translation helper is the one EventLinelist holds. This module used to
'build its own on every button press, and LLTranslation.Create validates all
'five translation tables.
Private Sub InitializeTrads()
    Dim linelistEvents As EventLinelist

    Set wb = ThisWorkbook
    Set linelistEvents = LinelistService()
    Set lltrads = linelistEvents.Translation

    'The 29 reads of the two translators below all assume a translator is there.
    'Building the helper here used to raise when the sheet was missing or a
    'table was broken, and this keeps that: the workbook is unusable either way,
    'and the raise says which workbook is at fault.
    If lltrads Is Nothing Then _
        Err.Raise ProjectError.ObjectNotInitialized, "EventsLinelistButtons", _
                  "This linelist carries no usable translation sheet"

    'Both scopes are cached inside the helper and answered on the first read.
    Set tradsmess = lltrads.TransObject()
    Set tradsform = lltrads.TransObject(TranslationOfForms)

    Set pass = linelistEvents.PasswordManager()
    Set wkbNames = linelistEvents.WorkbookNames()
End Sub

'Tell the user why a button refused to act.
Private Sub WarningOnSheet(ByVal msgCode As String)
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistService()
    linelistEvents.Warn msgCode
End Sub

'Tell the user a button failed. The detail carries the error description.
Private Sub FailureOnSheet(ByVal msgCode As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistService()
    linelistEvents.Fail msgCode, detail
End Sub

'The user log the event service holds. A workbook whose log cannot be
'built answers Nothing and every log line below stays quiet.
Private Function UserLogOf() As LLLog
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistService()
    Set UserLogOf = linelistEvents.UserLog()
End Function

'Write the success line of a finished walk. The write is guarded so a log
'fault never takes down the walk it records.
Private Sub LogSuccessLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSuccess action, detail
    On Error GoTo 0
End Sub

'Write the failure line of a walk that swallows its error at its label.
Private Sub LogFailureLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogFailure action, detail
    On Error GoTo 0
End Sub

'Write the warning line of a walk that ended with refused writes.
Private Sub LogWarningLine(ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogWarning action, detail
    On Error GoTo 0
End Sub

'The Err check of a swallowing handler label. The success path falls into
'the label with Err at 0 and gets its success line; the error jump gets the
'failure line. The caller reads Err into the two middle arguments before
'any cleanup call, because a called procedure's On Error statements clear
'Err.
Private Sub LogOutcomeLine(ByVal action As String, ByVal errNumber As Long, _
                           ByVal errDetail As String, _
                           Optional ByVal detail As String = vbNullString)
    If errNumber = 0 Then
        LogSuccessLine action, detail
    Else
        LogFailureLine action, errDetail
    End If
End Sub

'The show/hide line carries the writes Excel refused, read off the layout.
'A count above zero is a warning naming the sheet and the count, which is
'what surfaces a protected sheet or a position that is gone.
Private Sub LogShowHideLine(ByVal action As String, _
                            ByVal layout As ShowHideLayout, _
                            ByVal sheetName As String)
    If layout Is Nothing Then Exit Sub

    If layout.FailureCount > 0 Then
        LogWarningLine action, sheetName & ": " & layout.FailureCount & " refused writes"
    Else
        LogSuccessLine action, sheetName
    End If
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

'The hidden names of one worksheet.
'
'The service holds one store per sheet and drops it when that sheet raises a
'change, so the three readers below share one walk. A click on a sheet whose
'names cannot be read answers Nothing, and each reader then gives its empty
'value, which the callers already treat as "this is the wrong sheet".
Private Function SheetStoreOf(ByVal sh As Worksheet) As HiddenNames
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistService()
    Set SheetStoreOf = linelistEvents.SheetNames(sh)
End Function

'Get the sheet type tag.
Private Function SheetTag(ByVal sh As Worksheet) As String
    Dim shHn As HiddenNames

    Set shHn = SheetStoreOf(sh)
    If shHn Is Nothing Then Exit Function

    SheetTag = shHn.ValueAsString("sheet_type")
End Function

'Get the table name from worksheet-level HiddenNames.
Private Function TableNameOf(ByVal sh As Worksheet) As String
    Dim shHn As HiddenNames

    Set shHn = SheetStoreOf(sh)
    If shHn Is Nothing Then Exit Function

    TableNameOf = shHn.ValueAsString("table_name")
End Function

'The number of filled cells an untouched data row of this sheet carries.
'LLDataEntry writes it when it makes the table, and a row holding more filled
'cells than this is a row the user has typed into.
Private Function BlankRowCountOf(ByVal sh As Worksheet) As Long
    Dim shHn As HiddenNames

    Set shHn = SheetStoreOf(sh)
    If shHn Is Nothing Then Exit Function

    BlankRowCountOf = shHn.ValueAsLong("blank_row_count")
End Function

'The table name a sheet's PRINTSTART anchor is named after. VarWriter names the
'anchor after the base table name, and a printed sheet stores its own table name
'with the print_ prefix in front, so the prefix comes off here.
Private Function BaseTableNameOf(ByVal sh As Worksheet) As String
    Dim tabName As String

    tabName = TableNameOf(sh)
    If InStr(1, tabName, PRINTPREFIX, vbTextCompare) = 1 Then
        tabName = Mid$(tabName, Len(PRINTPREFIX) + 1)
    End If

    BaseTableNameOf = tabName
End Function

'The dictionary of the running linelist
Private Function DictionaryObject() As LLdictionary
    Set DictionaryObject = LLdictionary.Create(wb.Worksheets(DICTSHEET), 1, 1)
End Function

'The entry list of one sheet
Private Function EntriesFor(ByVal sh As Worksheet, _
                         ByVal layer As Byte, _
                         ByVal dict As LLdictionary) As ShowHide
    Set EntriesFor = ShowHide.Create(dict, layer, BaseSheetName(sh))
End Function

'The sheet half of the pair
Private Function LayoutFor(ByVal sh As Worksheet, ByVal layer As Byte) As ShowHideLayout
    Set LayoutFor = ShowHideLayout.Create(sh, layer, pass, BaseTableNameOf(sh))
End Function

'Build the entry list and the layout the form works on. Answers False on a sheet
'that has no show/hide layer.
Private Function OpenShowHideFor(ByVal sh As Worksheet) As Boolean
    Dim layer As Byte

    Set showHideEntries = Nothing
    Set showHideLayout = Nothing

    layer = ResolveShowHideLayer(SheetTag(sh))
    If layer = 0 Then Exit Function

    Set showHideEntries = EntriesFor(sh, layer, DictionaryObject())
    Set showHideLayout = LayoutFor(sh, layer)

    OpenShowHideFor = True
End Function

'The saved choices of this workbook. Answers Nothing when the workbook carries
'no show/hide worksheet, which is what an older linelist looks like.
Private Function ShowHideStoreOf() As ShowHideStore
    On Error Resume Next
    Set ShowHideStoreOf = ShowHideStore.Create(wb)
    On Error GoTo 0
End Function

'Read the saved choices of one entries back in. Answers how many rows matched.
Private Function LoadShowHideState(ByVal entries As ShowHide, _
                                  ByVal layout As ShowHideLayout) As Long
    Dim store As ShowHideStore

    If entries Is Nothing Then Exit Function

    Set store = ShowHideStoreOf()
    If store Is Nothing Then Exit Function

    LoadShowHideState = store.Load(entries, layout)
End Function

'Write the choices of one entries out. The layout is a parameter rather than the
'module one, because a caller that saves two layers in a row would otherwise
'record the printed sizes against the HList rows.
Private Sub SaveShowHideState(ByVal entries As ShowHide, ByVal layout As ShowHideLayout)
    Dim store As ShowHideStore

    If entries Is Nothing Then Exit Sub

    Set store = ShowHideStoreOf()
    If store Is Nothing Then Exit Sub

    store.Save entries, layout
End Sub

'Populate a show/hide form's list control from the entry list
Private Sub PopulateShowHideList(ByVal frm As Object)
    Dim listCtrl As Object
    Dim counter As Long

    If frm Is Nothing Then Exit Sub
    If showHideEntries Is Nothing Then Exit Sub

    On Error Resume Next
    If frm.Name = "F_ShowHideLL" Then
        Set listCtrl = frm.Controls("LST_LLVarNames")
    Else
        Set listCtrl = frm.Controls("LST_PrintNames")
    End If
    On Error GoTo 0

    If listCtrl Is Nothing Then Exit Sub

    listCtrl.Clear
    For counter = 1 To showHideEntries.EntryCount
        listCtrl.AddItem showHideEntries.HeaderText(counter)
    Next
End Sub

'@Description("Callback for click on show/hide in a linelist worksheet on a button")
'@EntryPoint
Public Sub ClickShowHide()
    Attribute ClickShowHide.VB_Description = "Callback for click on show/hide in a linelist worksheet on a button"

    Dim sh As Worksheet
    Dim shType As String
    Dim frm As Object

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    If (shType <> "HList" And shType <> "HList Print" And shType <> "VList" _
       And shType <> "HList CRF") Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    InitializeTrads

    If Not OpenShowHideFor(sh) Then Exit Sub

    'Read the choices of the last session and put the sheet in step with them.
    'With nothing saved yet, the sheet is the record: read it into the entries, so
    'a column the user hid by hand shows as hidden in the form.
    If LoadShowHideState(showHideEntries, showHideLayout) > 0 Then
        showHideEntries.Apply showHideLayout
    Else
        showHideEntries.Adopt showHideLayout
    End If

    'A CRF holds one variable per row and reads its labels straight, so it takes
    'the same form a data entry sheet does. The printed form is the only one
    'carrying the two header direction buttons.
    If shType = "HList Print" Then
        Set frm = F_ShowHidePrint
    Else
        Set frm = F_ShowHideLL
    End If

    Set activeShowHideForm = frm
    PopulateShowHideList frm
    frm.Show

    'After form closes, save the choices. The log line covers the whole form
    'session: every option click landed on this layout, so its refused-write
    'count is the count of the session.
    SaveShowHideState showHideEntries, showHideLayout
    LogShowHideLine "showhide", showHideLayout, sh.Name
    Set activeShowHideForm = Nothing
End Sub

'@Description("Callback for click on show/hide section in a linelist worksheet")
'@EntryPoint
Public Sub ClickShowHideSection()
    Attribute ClickShowHideSection.VB_Description = "Callback for click on show/hide section in a linelist worksheet"

    Dim sh As Worksheet
    Dim shType As String
    Dim layer As Byte
    Dim secMap As SectionMap
    Dim touched As BetterArray
    Dim firstPos As Long
    Dim lastPos As Long
    Dim hideThem As Boolean

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Sections are laid out on the data entry sheets alone. The printed
    'companion carries the same columns and the CRF its own rows, and
    'SectionBuilder writes a map for neither.
    If (shType <> "HList") And (shType <> "VList") Then
        WarningOnSheet "MSG_DataSheet"
        Exit Sub
    End If

    InitializeTrads

    If Not SelectionSpan(shType, firstPos, lastPos) Then
        WarningOnSheet "MSG_WrongCells"
        Exit Sub
    End If

    Set secMap = SectionMap.Create(sh)
    Set touched = secMap.IndicesInRange(firstPos, lastPos)

    If touched.Length = 0 Then
        WarningOnSheet "MSG_WrongCells"
        Exit Sub
    End If

    layer = ResolveShowHideLayer(shType)
    If layer = 0 Then Exit Sub

    Set showHideEntries = EntriesFor(sh, layer, DictionaryObject())
    Set showHideLayout = LayoutFor(sh, layer)

    'The same reconciliation ClickShowHide does. With nothing saved yet the
    'sheet is the record, so a column the user hid by hand is read as hidden
    'and the toggle below agrees with what the user can see.
    If LoadShowHideState(showHideEntries, showHideLayout) > 0 Then
        showHideEntries.Apply showHideLayout
    Else
        showHideEntries.Adopt showHideLayout
    End If

    'One press hides, the next shows. Every section the selection touches has
    'to be hidden already before the press is read as "show them again", so a
    'selection spanning a hidden section and a visible one hides both rather
    'than half.
    hideThem = Not AllSectionsHidden(secMap, touched)

    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState
    ApplyToSections secMap, touched, hideThem
    showHideEntries.Apply showHideLayout
    SaveShowHideState showHideEntries, showHideLayout
    LogShowHideLine "showhide-section", showHideLayout, sh.Name

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "showhide-section", Err.Description
    LinelistEventsManager.LLExitBusyState
    Set showHideEntries = Nothing
    Set showHideLayout = Nothing
End Sub

'The span of positions the user has selected, in the axis the sheet hides on:
'columns on an HList sheet, rows on a VList one. Answers True for a cell
'selection. A chart or a shape holds the selection as an object of its own, and
'those answer False.
Private Function SelectionSpan(ByVal shType As String, _
                              ByRef firstPos As Long, _
                              ByRef lastPos As Long) As Boolean
    Dim rng As Range

    firstPos = 0
    lastPos = 0

    If TypeName(Application.Selection) <> "Range" Then Exit Function
    Set rng = Application.Selection

    If shType = "VList" Then
        firstPos = rng.Row
        lastPos = firstPos + rng.Rows.Count - 1
    Else
        firstPos = rng.Column
        lastPos = firstPos + rng.Columns.Count - 1
    End If

    SelectionSpan = (firstPos > 0)
End Function

'Whether every section of the list is already hidden. A section holding nothing
'the user owns is passed over: it can never be hidden, so counting it would
'leave a selection that touches one stuck on "hide" for good.
Private Function AllSectionsHidden(ByVal secMap As SectionMap, _
                                  ByVal touched As BetterArray) As Boolean
    Dim counter As Long
    Dim blockIdx As Long
    Dim state As Byte
    Dim answered As Long

    If showHideEntries Is Nothing Then Exit Function

    For counter = touched.LowerBound To touched.UpperBound
        blockIdx = CLng(touched.Item(counter))
        state = showHideEntries.RangeState(secMap.StartAt(blockIdx), _
                                           secMap.EndAt(blockIdx))

        Select Case state
        Case ShowHideRangeEmpty, ShowHideRangeFixed
            'nothing to say either way
        Case ShowHideRangeHidden
            answered = answered + 1
        Case Else
            Exit Function
        End Select
    Next counter

    AllSectionsHidden = (answered > 0)
End Function

'Hide or show every section of the list.
Private Sub ApplyToSections(ByVal secMap As SectionMap, _
                           ByVal touched As BetterArray, _
                           ByVal hidden As Boolean)
    Dim counter As Long
    Dim blockIdx As Long

    If showHideEntries Is Nothing Then Exit Sub

    For counter = touched.LowerBound To touched.UpperBound
        blockIdx = CLng(touched.Item(counter))
        showHideEntries.SetHiddenInRange secMap.StartAt(blockIdx), _
                                         secMap.EndAt(blockIdx), _
                                         hidden
    Next counter
End Sub

'@Description("Callback for click on the list of showhide")
'@EntryPoint
Public Sub ClickListShowHide(ByVal Index As Long)
    Attribute ClickListShowHide.VB_Description = "Callback for click on the list of showhide"

    Dim entryIdx As Long
    Dim canChange As Boolean
    Dim isVertical As Boolean

    If showHideEntries Is Nothing Then Exit Sub
    If activeShowHideForm Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideEntries.EntryCount Then Exit Sub

    'A mandatory entry and a locked entry both refuse the user, so the one
    'question the buttons ask is whether the entry is free.
    canChange = showHideEntries.IsFree(entryIdx)

    If activeShowHideForm.Name = "F_ShowHideLL" Then
        If showHideEntries.IsHidden(entryIdx) Then
            activeShowHideForm.OPT_Hide.Value = True
        Else
            activeShowHideForm.OPT_Show.Value = True
        End If
        activeShowHideForm.OPT_Show.Enabled = canChange
        activeShowHideForm.OPT_Hide.Enabled = canChange
    Else
        isVertical = False
        If Not showHideLayout Is Nothing Then
            isVertical = showHideLayout.IsVertical(showHideEntries.PositionIndex(entryIdx))
        End If

        If showHideEntries.IsHidden(entryIdx) Then
            activeShowHideForm.OPT_Hide.Value = True
        ElseIf isVertical Then
            activeShowHideForm.OPT_PrintShowVerti.Value = True
        Else
            activeShowHideForm.OPT_PrintShowHoriz.Value = True
        End If
        activeShowHideForm.OPT_PrintShowHoriz.Enabled = canChange
        activeShowHideForm.OPT_PrintShowVerti.Enabled = canChange
        activeShowHideForm.OPT_Hide.Enabled = canChange
    End If
End Sub

'@Description("Callback for clik on differents show hide options on a button")
'@EntryPoint
Public Sub ClickOptionsShowHide(ByVal Index As Long)
    Attribute ClickOptionsShowHide.VB_Description = "Callback for clik on differents show hide options on a button"

    Dim entryIdx As Long
    Dim shouldHide As Boolean
    Dim wantsVertical As Boolean
    Dim posIdx As Long

    If showHideEntries Is Nothing Then Exit Sub
    If showHideLayout Is Nothing Then Exit Sub
    If activeShowHideForm Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideEntries.EntryCount Then Exit Sub
    If Not showHideEntries.IsFree(entryIdx) Then Exit Sub

    shouldHide = activeShowHideForm.OPT_Hide.Value
    showHideEntries.SetHidden entryIdx, shouldHide

    posIdx = showHideEntries.PositionIndex(entryIdx)
    If posIdx = 0 Then Exit Sub

    showHideLayout.BeginBatch
    showHideLayout.SetHidden posIdx, shouldHide

    'The header direction belongs to the two show buttons of the printed form.
    'Choosing Hide leaves the direction as it was, so showing the column again
    'brings back the header the user had.
    If showHideLayout.SupportsOrientation And Not shouldHide Then
        wantsVertical = False
        On Error Resume Next
        wantsVertical = activeShowHideForm.OPT_PrintShowVerti.Value
        On Error GoTo 0
        showHideLayout.SetOrientation posIdx, wantsVertical
    End If

    showHideLayout.EndBatch
End Sub

'@Description("Callback for click on column width in show/hide")
'@EntryPoint
Public Sub ClickColWidth(ByVal Index As Long)
    Attribute ClickColWidth.VB_Description = "Callback for click on column width in show/hide"

    Dim entryIdx As Long
    Dim posIdx As Long
    Dim inputValue As String

    If showHideEntries Is Nothing Then Exit Sub
    If showHideLayout Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideEntries.EntryCount Then Exit Sub

    posIdx = showHideEntries.PositionIndex(entryIdx)
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

    showHideLayout.SetSize posIdx, CDbl(inputValue)
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
    LogOutcomeLine "open-print", Err.Number, Err.Description
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
    LogOutcomeLine "open-crf", Err.Number, Err.Description
    pass.Protect wb
End Sub

'@Description("Callback for click on close print, CRF or log sheet")
'@EntryPoint
Public Sub ClickCloseSheet()
    Attribute ClickCloseSheet.VB_Description = "Callback for click on close print/crf/log sheet"

    Dim sh As Worksheet
    Dim shType As String
    Dim printsh As Worksheet
    Dim crfsh As Worksheet
    Dim logSheetName As String
    Dim actionCode As String

    On Error GoTo ErrClose
    actionCode = "close-print"
    Set sh = ActiveSheet

    InitializeTrads

    shType = SheetTag(sh)

    'The log sheet carries no sheet tag; its name is the mark.
    logSheetName = LLLog.SheetName

    If shType <> "HList" And shType <> "HList Print" And shType <> "HList CRF" _
       And sh.Name <> logSheetName Then
        WarningOnSheet "MSG_PrintCRFOrDataSheet"
        Exit Sub
    End If

    'Unprotect workbook
    pass.UnProtect wb

    If sh.Name = logSheetName Then
        actionCode = "close-log"
        sh.Visible = xlSheetVeryHidden
    ElseIf shType = "HList" Then
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
    LogOutcomeLine actionCode, Err.Number, Err.Description
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
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow

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
    LogOutcomeLine "rotate", Err.Number, Err.Description
    LinelistEventsManager.LLExitBusyState
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
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow

    If shType = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)
    pass.UnProtect sh.Name

    Set Lo = sh.ListObjects(1)
    If (Lo.DataBodyRange Is Nothing) Then
        Set LoRng = Lo.HeaderRowRange.Offset(1)
    Else
        Set LoRng = Lo.DataBodyRange
    End If

    'Ask for rowheight. A plain Exit Sub here would skip the busy-state exit
    'and leave the application locked, so both refusals go through ErrHand.
    Do While (True)
        inputValue = InputBox(tradsmess.TranslatedValue("MSG_RowHeight"), _
                             tradsmess.TranslatedValue("MSG_Enter"))
        If inputValue = vbNullString Then GoTo ErrHand
        If IsNumeric(inputValue) Then Exit Do
        If (MsgBox(tradsmess.TranslatedValue("MSG_EnterNumeric"), _
             vbOkCancel, vbNullString) = vbCancel) Then GoTo ErrHand
    Loop

    On Error Resume Next
        actualRowHeight = CLng(inputValue)
        LoRng.EntireRow.RowHeight = actualRowHeight
    On Error GoTo 0

    'The success line sits above the label so a cancelled ask logs nothing.
    LogSuccessLine "row-height", inputValue

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "row-height", Err.Description
    LinelistEventsManager.LLExitBusyState
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
        LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow
        'Unprotect current worksheet
        pass.UnProtect "_active"
        'remove the filters
        Lo.AutoFilter.ShowAllData
        pass.Protect "_active"
        LinelistEventsManager.LLExitBusyState
        LogSuccessLine "remove-filters", sh.Name
    End If
    Exit Sub

ErrHand:
    LogFailureLine "remove-filters", Err.Description
    pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
End Sub

'@Description("Add rows to a data entry table in the Linelist")
'@EntryPoint
Public Sub ClickAddRows()
    Attribute ClickAddRows.VB_Description = "Add rows to a data entry table in the Linelist"

    Dim Lo As listObject
    Dim csTab As CustomTable
    Dim sh As Worksheet
    Dim shType As String
    Dim nbRows As Long

    On Error GoTo errAddRows
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow
    InitializeTrads

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet. The busy state is already on,
    'so the refusal leaves through the cleanup label below.
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        GoTo Cleanup
    End If

    'The busy state keeps events off while the rows come in.
    pass.UnProtect "_active"

    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)
    nbRows = IIf(shType = "HList", 199, 10)
    csTab.AddRows nbRows:=nbRows
    LogSuccessLine "add-rows", nbRows & " rows on " & sh.Name

Cleanup:
    'Protect only HList
    If shType = "HList" Then pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    Exit Sub

errAddRows:
    On Error Resume Next
    If shType = "HList" Then pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    FailureOnSheet "MSG_ErrAddRows"
    On Error GoTo 0
End Sub

'@Description("Resize the data entry table in the linelist")
'@EntryPoint
Public Sub ClickResize()
    Attribute ClickResize.VB_Description = "Resize the data entry table in the linelist"

    Dim Lo As listObject
    Dim csTab As CustomTable
    Dim sh As Worksheet
    Dim shType As String
    Dim nbBlank As Long

    On Error GoTo errDelRows
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlWait
    InitializeTrads

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet. The busy state is already on,
    'so the refusal leaves through the cleanup label below.
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "MSG_PrintOrDataSheet"
        GoTo Cleanup
    End If

    'The busy state keeps events off while the rows go out.
    pass.UnProtect "_active"

    nbBlank = BlankRowCountOf(sh)
    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)

    csTab.RemoveRows totalCount:=nbBlank
    LogSuccessLine "resize", sh.Name

Cleanup:
    If shType = "HList" Then pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    Exit Sub

errDelRows:
    On Error Resume Next
    If shType = "HList" Then pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    FailureOnSheet "MSG_ErrDelRows"
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

    Dim topPosition As Single
    Dim expObj As LLExport
    Dim expsh As Worksheet

    'initialize translations
    InitializeTrads

    Set expsh = ThisWorkbook.Worksheets(EXPORTSHEET)
    Set expObj = LLExport.Create(expsh)

    On Error GoTo errLoadExp

    'The form owns its export buttons: it adds one per active export, it holds
    'the object that answers each click, and it takes them off again below.
    'The qualified call runs the F_Export code-behind, whose source is the
    'FormLogicExport module. It answers the free position under the last
    'button, which the fixed controls are laid out from.
    topPosition = F_Export.SetupExportForm(ThisWorkbook, tradsmess, expObj, _
                                           COMMANDGAPS, COMMANDHEIGHT, COMMANDGAPS)

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

    'The form takes its own buttons off on both ways out, and SetupExportForm
    'clears whatever is left before it adds any. Reading F_Export here would
    'build the instance back up when the user closed it with the window box.
    LogSuccessLine "export"

    Exit Sub

errLoadExp:
    On Error Resume Next
    F_Export.TeardownExportForm
    FailureOnSheet "MSG_ErrLoadExport"
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
                LoadGeo GeoScopeAdmin
            Case "hf"
                LoadGeo GeoScopeHF
            Case Else
                WarningOnSheet "MSG_WrongCells"
            End Select
        Else
            WarningOnSheet "MSG_WrongCells"
        End If

    Case "SPT-Analysis"
        On Error Resume Next
        rngName = ActiveCell.Name.Name
        On Error GoTo 0
        If (InStr(1, rngName, "INPUTSPTGEO_") > 0) Then
            LoadGeo GeoScopeAdmin
        ElseIf (InStr(1, rngName, "INPUTSPTHF_") > 0) Then
            LoadGeo GeoScopeHF
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
    Dim counter As Long

    'A guard used to sit here to skip a second click that came soon after the
    'first. It never once skipped anything: it read
    '"If (timePeriod < 60 & timePeriod > 0)", where & joins two strings and binds
    'tighter than the comparisons, so the test was always False. Its clock was
    'wrong too -- a Long held a whole day number and was written on the first
    'click alone. Calculate is a button the user presses on purpose, so the
    'answer is to calculate every time, which is what the workbook has always
    'done.
    InitializeTrads
    On Error GoTo ErrHand

    'Calculate. UpdateSpTables enters the same busy state underneath;
    'busyDepth counts the nesting and the restore happens here alone.
    LinelistEventsManager.LLEnterBusyState
    UpdateSpTables

    Set anaSheetsList = New BetterArray
    anaSheetsList.Push wkbNames.ValueAsString("RNG_UASheet"), _
                       wkbNames.ValueAsString("RNG_TSSheet"), _
                       wkbNames.ValueAsString("RNG_SPSheet"), _
                       wkbNames.ValueAsString("RNG_SPTSheet")

    For counter = anaSheetsList.LowerBound To anaSheetsList.UpperBound
        sheetName = anaSheetsList.Item(counter)
        Set sh = wb.Worksheets(sheetName)
        'One pass covers the sheet. These four sheets carry array formulas over
        'the whole linelist, and the second pass over columns A to E was inside
        'the first. HandleAnalysisChange calculates the same sheets this way.
        sh.calculate
    Next

ErrHand:
    LogOutcomeLine "calculate", Err.Number, Err.Description
    LinelistEventsManager.LLExitBusyState
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

    Dim linelistEvents As EventLinelist
    Dim varLabTab As BetterArray

    InitializeTrads
    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState

    'The rows come from the held dictionary and variable reader. This used to
    'build both per click and stage the rows on the __temp worksheet.
    Set linelistEvents = LinelistService()
    Set varLabTab = linelistEvents.VarLabelTable()

    'Affect the table to the list
    F_ShowVarLabels.LST_CustomTabList.List = varLabTab.Items
    LinelistEventsManager.LLExitBusyState

    'This will open the form with variable name and variable labels for
    [F_ShowVarLabels].Show
    Exit Sub

ErrHand:
    LogFailureLine "open-varlab", Err.Description
    LinelistEventsManager.LLExitBusyState
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
        
        LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow
        'Unprotect the active worksheet, sort the range, and protect back.
        'I have to keep the protect/unprotect step as far as possible for
        'performance issues
        pass.UnProtect "_active"
        On Error Resume Next
        LoRng.Sort key1:=sortRng, Order1:=sortOrder, Header:=xlYes
        On Error GoTo 0
        pass.Protect "_active"
        LinelistEventsManager.LLExitBusyState
        LogSuccessLine "sort", headerName
    End If
    Exit Sub

ErrHand:
    LogFailureLine "sort", Err.Description
    LinelistEventsManager.LLExitBusyState
End Sub


'@Description("Export all Analysis worksheets to a workbook")
'@EntryPoint
Public Sub ClickExportAnalysis()
    Attribute ClickExportAnalysis.VB_Description = "Export all Analysis worksheets to a workbook"

    InitializeTrads

    'The walk lives in the F_ExportMig code-behind, whose source is the
    'FormLogicExportMig module. The qualified call runs the form's copy; the
    'standard-module copy carries form references and is never compiled.
    F_ExportMig.HandleAnalysisExport wb, tradsmess
End Sub

'@Description("Import new data in the linelist")
'@EntryPoint
Public Sub ClickImportData()
    Attribute clickImportData.VB_Description = "Import new data in the linelist"

    Dim sh As Worksheet
    Dim csTab As CustomTable
    Dim Lo As ListObject
    Dim nbBlank As Long

    InitializeTrads

    'Resize all HList tables before import (remove blank rows)
    Application.EnableEvents = False
    For Each sh In wb.Worksheets
        If SheetTag(sh) = "HList" Then
            nbBlank = BlankRowCountOf(sh)
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

    'Import data using LLImporter API (handles file picker, the question about
    'the rows already entered, busy state and report). Qualified through the
    'form: the walk lives in the F_Advanced code-behind.
    F_Advanced.HandleImportData wb, tradsmess

    'Update all the listAuto in the workbook
    LinelistEventsManager.UpdateAllListAuto

    Application.EnableEvents = True
End Sub

'@Description("Import a new geobase in the linelist")
'@EntryPoint
Public Sub ClickImportGeobase()
    Attribute clickImportGeobase.VB_Description = "Import a new geobase in the linelist"

    InitializeTrads
    F_Advanced.HandleImportGeobase wb, tradsmess
End Sub

'@Description("Reset hidden columns in the linelist")
'@EntryPoint
Public Sub ClickResetColumns()
    Attribute clickResetColumns.VB_Description = "Reset hidden columns in the linelist"

    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState
    InitializeTrads

    'The walk lives in the F_Advanced code-behind. The Reset Columns button
    'of F_Advanced comes through here for the busy state, and the walk is
    'reached back through the form.
    F_Advanced.HandleResetColumns wb, pass

ErrHand:
    LogOutcomeLine "reset-columns", Err.Number, Err.Description
    LinelistEventsManager.LLExitBusyState
End Sub

'@Description("Hide/Unhide Optional variables in the linelist")
'@EntryPoint
Public Sub ClickShowHideMinimal()
    Attribute ClickShowHideMinimal.VB_Description = "Hide/Unhide Optional variables in the linelist"

    Const RNGSHOWALLOPTIONALS As String = "RNG_ShowAllOptionals"

    Dim showOptional As Boolean
    Dim checkConfirm As Boolean
    Dim sh As Worksheet

    On Error GoTo ErrHand

    LinelistEventsManager.LLEnterBusyState
    InitializeTrads

    showOptional = (wkbNames.ValueAsString(RNGSHOWALLOPTIONALS) = "yes")

    'Warn user: they will lose current shown/hidden status
    checkConfirm = (MsgBox(tradsmess.TranslatedValue("MSGB_WarningShowHide"), _
                           vbYesNo + vbExclamation, _
                           tradsmess.TranslatedValue("MSGB_Warning")) = vbYes)

    If Not checkConfirm Then GoTo ErrHand

    Set sh = ActiveSheet
    If Not OpenShowHideFor(sh) Then GoTo ErrHand

    'Toggle every entry the user owns, then put the sheet in step
    showHideEntries.SetAllOptionalHidden Not showOptional
    showHideEntries.Apply showHideLayout

    'Toggle the flag
    If showOptional Then
        wkbNames.SetValue RNGSHOWALLOPTIONALS, "no"
    Else
        wkbNames.SetValue RNGSHOWALLOPTIONALS, "yes"
    End If

    SaveShowHideState showHideEntries, showHideLayout

    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If

    'The success line sits above the label so a declined confirm logs nothing.
    LogShowHideLine "showhide-minimal", showHideLayout, sh.Name

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "showhide-minimal", Err.Description
    LinelistEventsManager.LLExitBusyState
End Sub



'@Description("Callback for click on the layouts button of the show/hide form")
'@EntryPoint
Public Sub ClickShowHideLayouts()
    Attribute ClickShowHideLayouts.VB_Description = "Callback for click on the layouts button of the show/hide form"

    F_ShowHideSave.Show

    'The show/hide form hides itself before this runs, but ClickShowHide still
    'saves the pair once the click handler returns. The pair is read back from
    'the store here, so that save records what the sheets now show rather than
    'the choices from before a restore.
    If showHideEntries Is Nothing Then Exit Sub

    If LoadShowHideState(showHideEntries, showHideLayout) > 0 Then
        showHideEntries.Apply showHideLayout
    Else
        showHideEntries.Adopt showHideLayout
    End If

    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If
End Sub

'@Description("Match the show/hide state in the linelist from the print sheet")
'@EntryPoint
Public Sub ClickMatchLinelistShowHide()
    Attribute ClickMatchLinelist.VB_Description = "Match the show/hide state in the linelist from the print sheet"

    Dim checkConfirm As Boolean
    Dim baseEntries As ShowHide
    Dim printEntries As ShowHide
    Dim printLayout As ShowHideLayout
    Dim printsh As Worksheet
    Dim sh As Worksheet
    Dim dict As LLdictionary
    Dim sheetName As String
    Dim counter As Long
    Dim printIdx As Long
    Dim posIdx As Long

    On Error GoTo ErrHand

    LinelistEventsManager.LLEnterBusyState
    InitializeTrads

    'Warn user: they will lose current shown/hidden status
    checkConfirm = (MsgBox(tradsmess.TranslatedValue("MSGB_WarningShowHide"), _
                           vbYesNo + vbExclamation, _
                           tradsmess.TranslatedValue("MSGB_Warning")) = vbYes)

    If Not checkConfirm Then GoTo ErrHand

    Set printsh = ActiveSheet
    sheetName = Mid$(printsh.Name, Len(PRINTPREFIX) + 1)
    Set sh = wb.Worksheets(sheetName)

    Set dict = DictionaryObject()
    Set baseEntries = ShowHide.Create(dict, ShowHideLayerHList, sheetName)
    Set printEntries = ShowHide.Create(dict, ShowHideLayerPrinted, sheetName)
    Set printLayout = LayoutFor(printsh, ShowHideLayerPrinted)

    'Read what the data entry sheet shows today, then copy it over
    LoadShowHideState baseEntries, Nothing
    baseEntries.Adopt LayoutFor(sh, ShowHideLayerHList)

    For counter = 1 To baseEntries.EntryCount
        printIdx = printEntries.IndexOf(baseEntries.FieldKey(counter))
        If printIdx > 0 Then printEntries.SetHidden printIdx, baseEntries.IsHidden(counter)
    Next

    printEntries.Apply printLayout

    'Every column the print sheet now shows reads horizontally
    printLayout.BeginBatch
    For counter = 1 To printEntries.EntryCount
        posIdx = printEntries.PositionIndex(counter)
        If posIdx > 0 And Not printEntries.IsHidden(counter) Then
            printLayout.SetOrientation posIdx, False
        End If
    Next
    printLayout.EndBatch

    SaveShowHideState baseEntries, Nothing
    SaveShowHideState printEntries, printLayout

    'Sync the module state so the open form reads the list that was just saved
    Set showHideEntries = printEntries
    Set showHideLayout = printLayout
    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If

    'The success line sits above the label so a declined confirm logs nothing.
    LogShowHideLine "showhide-match", printLayout, printsh.Name

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "showhide-match", Err.Description
    LinelistEventsManager.LLExitBusyState
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

    LinelistEventsManager.LLEnterBusyState
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
    'Err is read before the Resume Next below clears it.
    LogOutcomeLine "autofit", Err.Number, Err.Description, sh.Name
    On Error Resume Next
    LinelistEventsManager.LLExitBusyState
End Sub
