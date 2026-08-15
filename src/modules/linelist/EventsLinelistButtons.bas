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
'The suffix a sheet's hidden name store puts on a variable name to key its
'dictionary control. Spelled the same way in LLSpatial, and the spacing is part
'of the key, so it is copied here exactly rather than rebuilt.
Private Const CONTROL_SUFFIX As String = " -- control"

'The translation code the sections list shows for a block of variables the
'dictionary gave no `main section`. It sits in the forms table beside the two
'the status column reads, so the whole list speaks one language.
Private Const EMPTY_SECTION_TAG As String = "LBL_EmptySection"

'The pair the show/hide form is open on: which variables the sheet offers, and
'the sheet itself. Both are rebuilt each time the form opens.
'
'activeLayout, NOT showHideLayout. VBA is not case sensitive, so a module-level
'`showHideLayout` and the class `ShowHideLayout` are one identifier, and the
'variable wins everywhere in this module. LayoutFor below then reads
'`ShowHideLayout.Create` as a call on the variable, which is Nothing until
'LayoutFor returns -- so every press of the show/hide button raised 91, "Object
'variable or With block variable not set", on a line that named the class.
'scripts/devtools/name-shadowing-scan.R exists to catch the next one.
Private showHideEntries As ShowHide
Private activeLayout As ShowHideLayout
Private activeShowHideForm As Object
Private tradsform As TranslationObject   'Translation of forms
Private tradsmess As TranslationObject   'Translation of messages
Private pass As Passwords
Private wb As Workbook
Private lltrads As LLTranslation
Private wkbNames As HiddenNames

'The pair the sections form is open on: the sections of the sheet, and the form
'itself. Both are rebuilt each time the form opens.
Private activeSections As SectionShowHide
Private activeSectionsForm As Object

'The section the last press of the section button collapsed, and the sheet it
'was collapsed on. The press that follows brings that section back.
'
'Hiding a section collapses its columns, and Excel moves the cursor to the next
'visible column when it does. That column is inside the NEXT section, so a
'button that reads the cursor a second time acts on a section the user never
'pointed at. The cursor is parked off the title line after a hide and the
'section is remembered here instead.
Private lastHiddenSectionSheet As String
Private lastHiddenSectionIndex As Long

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

'Every log line of this module names the procedure that wrote it.
'
'VBA carries no call stack a helper could read a caller's name from, so the
'caller names itself and the name is the first argument of every routine
'below. It is required rather than optional on purpose: a call site that
'forgets it stops the module compiling, which is the only way to keep the
'names right as buttons are added.
'
'Tell the user why a button refused to act.
Private Sub WarningOnSheet(ByVal source As String, ByVal msgCode As String)
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistService()
    linelistEvents.Warn msgCode, source
End Sub

'Tell the user a button failed. The detail carries the error description and
'shows in the box behind the message; a reason meant for the log alone goes
'in logDetail, which keeps a raw VBA description out of a field user's way.
Private Sub FailureOnSheet(ByVal source As String, ByVal msgCode As String, _
                           Optional ByVal detail As String = vbNullString, _
                           Optional ByVal logDetail As String = vbNullString)
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistService()
    linelistEvents.Fail msgCode, detail, vbNullString, source, logDetail
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
Private Sub LogSuccessLine(ByVal source As String, ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogSuccess action, detail, source
    On Error GoTo 0
End Sub

'Write the failure line of a walk that swallows its error at its label.
Private Sub LogFailureLine(ByVal source As String, ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogFailure action, detail, source
    On Error GoTo 0
End Sub

'Write the warning line of a walk that ended with refused writes.
Private Sub LogWarningLine(ByVal source As String, ByVal action As String, _
                           Optional ByVal detail As String = vbNullString)
    Dim logStore As LLLog

    Set logStore = UserLogOf()
    If logStore Is Nothing Then Exit Sub

    On Error Resume Next
    logStore.LogWarning action, detail, source
    On Error GoTo 0
End Sub

'The Err check of a swallowing handler label. The success path falls into
'the label with Err at 0 and gets its success line; the error jump gets the
'failure line. The caller reads Err into the two middle arguments before
'any cleanup call, because a called procedure's On Error statements clear
'Err.
'The detail rides on both lines. It used to be dropped on the error path,
'which is the path that wants it most: a walk that failed halfway names the
'reason and says nothing about how far it got or which sheet it was on.
Private Sub LogOutcomeLine(ByVal source As String, ByVal action As String, _
                           ByVal errNumber As Long, _
                           ByVal errDetail As String, _
                           Optional ByVal detail As String = vbNullString)
    If errNumber = 0 Then
        LogSuccessLine source, action, detail
        Exit Sub
    End If

    If LenB(Trim$(detail)) = 0 Then
        LogFailureLine source, action, errDetail
    Else
        LogFailureLine source, action, errDetail & " (" & detail & ")"
    End If
End Sub

'The show/hide line carries the writes Excel refused, read off the layout.
'A count above zero is a warning naming what moved and the count, which is
'what surfaces a protected sheet or a position that is gone.
'
'The detail is built by the caller and says what actually moved -- which
'section, on which sheet, and which way. It used to be the sheet name alone,
'so seven presses on seven sections wrote seven identical lines.
Private Sub LogShowHideLine(ByVal source As String, ByVal action As String, _
                            ByVal layout As ShowHideLayout, _
                            ByVal detail As String)
    If layout Is Nothing Then Exit Sub

    If layout.FailureCount > 0 Then
        LogWarningLine source, action, _
                       detail & ", " & layout.FailureCount & " refused writes"
    Else
        LogSuccessLine source, action, detail
    End If
End Sub

'What one section press did, in the words the log wants: the section by name,
'which way it went, and the sheet it is on.
'
'SectionDisplayName is what the sections form lists, so a run of variables the
'dictionary left with no main section reads the same in both places instead of
'reaching the log as an empty name.
Private Function SectionMoveText(ByVal sections As SectionShowHide, _
                                 ByVal sectionIdx As Long, _
                                 ByVal hideIt As Boolean, _
                                 ByVal sheetName As String) As String
    Dim movedWay As String

    movedWay = IIf(hideIt, "hidden", "shown")

    If sections Is Nothing Then
        SectionMoveText = "section " & sectionIdx & " " & movedWay & _
                          " on " & sheetName
        Exit Function
    End If

    SectionMoveText = SectionDisplayName(sections.SectionNameAt(sectionIdx)) & _
                      " " & movedWay & " on " & sheetName
End Function

'How much of a sheet a show/hide session left standing. The form moves any
'number of variables before it closes, so a count is what its one log line
'can honestly say.
Private Function EntryCountText(ByVal entries As ShowHide, _
                                ByVal sheetName As String) As String
    Dim counter As Long
    Dim hiddenCount As Long

    If entries Is Nothing Then
        EntryCountText = sheetName
        Exit Function
    End If

    For counter = 1 To entries.EntryCount
        If entries.IsHidden(counter) Then hiddenCount = hiddenCount + 1
    Next counter

    EntryCountText = hiddenCount & " of " & entries.EntryCount & _
                     " variables hidden on " & sheetName
End Function

'How much of a sheet the sections form left standing. The form moves any
'number of sections in one session, so the count is what its log line can
'honestly say; naming them one by one would be a line per click.
Private Function SectionCountText(ByVal sections As SectionShowHide, _
                                  ByVal sheetName As String) As String
    Dim counter As Long
    Dim hiddenCount As Long

    If sections Is Nothing Then
        SectionCountText = sheetName
        Exit Function
    End If

    For counter = 1 To sections.Count
        If sections.CanChange(counter) Then
            If sections.IsHidden(counter) Then hiddenCount = hiddenCount + 1
        End If
    Next counter

    SectionCountText = hiddenCount & " of " & sections.Count & _
                       " sections hidden on " & sheetName
End Function

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
'The dictionary control of the column the cursor is in, in lower case.
'
'The variable NAMES of an HList sheet live on the row the table's _START anchor
'is on, and every one of them has a `<name> -- control` entry in the sheet's
'hidden name store. That pair is how LLSpatial reads a column's control, and it
'is the only reading that survives a change to the header layout.
'
'ClickGeoApp used to read Cells(startRow - 5, column) instead. With _START on
'row 8 and the first data row on row 9, that lands on row 4 -- four rows above
'the names -- so the answer was whatever happened to be there and never "geo1".
'The geo form refused every column, including the adm1 column it was standing
'on.
Private Function ColumnControl(ByVal sh As Worksheet, _
                               ByVal tabName As String, _
                               ByVal targetColumn As Long) As String
    Dim anchor As Range
    Dim shHn As HiddenNames
    Dim varName As String

    On Error Resume Next
        Set anchor = sh.Range(tabName & "_START")
    On Error GoTo 0
    If anchor Is Nothing Then Exit Function

    varName = Trim$(CStr(sh.Cells(anchor.Row, targetColumn).Value))
    If LenB(varName) = 0 Then Exit Function

    Set shHn = SheetStoreOf(sh)
    If shHn Is Nothing Then Exit Function

    ColumnControl = LCase$(shHn.ValueAsString(varName & CONTROL_SUFFIX))
End Function

Private Function LayoutFor(ByVal sh As Worksheet, ByVal layer As Byte) As ShowHideLayout
    Set LayoutFor = ShowHideLayout.Create(sh, layer, pass, BaseTableNameOf(sh))
End Function

'Build the entry list and the layout the form works on. Answers False on a sheet
'that has no show/hide layer.
Private Function OpenShowHideFor(ByVal sh As Worksheet) As Boolean
    Dim layer As Byte

    Set showHideEntries = Nothing
    Set activeLayout = Nothing

    layer = ResolveShowHideLayer(SheetTag(sh))
    If layer = 0 Then Exit Function

    Set showHideEntries = EntriesFor(sh, layer, DictionaryObject())
    Set activeLayout = LayoutFor(sh, layer)

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

'The list control of one show/hide form
Private Function ShowHideListOf(ByVal frm As Object) As Object
    If frm Is Nothing Then Exit Function

    On Error Resume Next
    If frm.Name = "F_ShowHideLL" Then
        Set ShowHideListOf = frm.Controls("LST_LLVarNames")
    Else
        Set ShowHideListOf = frm.Controls("LST_PrintNames")
    End If
    On Error GoTo 0
End Function

'The word the status column shows for one entry. The two words are the captions
'of the form's own option buttons, so the column reads in the language the rest
'of the form reads in. A form opened before the translators are built falls back
'to the tag itself.
Private Function ShowHideStatusText(ByVal hidden As Boolean) As String
    Dim tagName As String

    tagName = IIf(hidden, "OPT_Hide", "OPT_Show")
    ShowHideStatusText = tagName

    If tradsform Is Nothing Then Exit Function

    On Error Resume Next
    ShowHideStatusText = tradsform.TranslatedValue(tagName)
    On Error GoTo 0
End Function

'Populate a show/hide form's list control from the entry list
'
'The list carries three columns: the label the user reads, the variable name the
'dictionary spells, and whether the entry is shown or hidden right now. Only the
'first column was ever written, so the two beside it stayed blank on every open.
'ColumnCount is set before the rows go in, because writing List(row, 1) on a one
'column control is refused.
Private Sub PopulateShowHideList(ByVal frm As Object)
    Dim listCtrl As Object
    Dim counter As Long
    Dim rowIdx As Long
    Dim shownText As String
    Dim hiddenText As String

    If showHideEntries Is Nothing Then Exit Sub

    Set listCtrl = ShowHideListOf(frm)
    If listCtrl Is Nothing Then Exit Sub

    shownText = ShowHideStatusText(False)
    hiddenText = ShowHideStatusText(True)

    On Error Resume Next
    listCtrl.ColumnCount = 3
    On Error GoTo 0

    listCtrl.Clear
    For counter = 1 To showHideEntries.EntryCount
        rowIdx = counter - 1
        listCtrl.AddItem showHideEntries.HeaderText(counter)

        'A control the deployed form still carries as one column refuses these
        'two writes, and the label column is what the user needs most.
        On Error Resume Next
        listCtrl.List(rowIdx, 1) = showHideEntries.FieldKey(counter)
        listCtrl.List(rowIdx, 2) = IIf(showHideEntries.IsHidden(counter), _
                                       hiddenText, shownText)
        On Error GoTo 0
    Next
End Sub

'Rewrite the status cell of one row, after the user changed that entry. The
'whole list is left alone: rebuilding it would drop the selection the user is
'working from.
Private Sub RefreshShowHideRow(ByVal entryIdx As Long)
    Dim listCtrl As Object

    If showHideEntries Is Nothing Then Exit Sub

    Set listCtrl = ShowHideListOf(activeShowHideForm)
    If listCtrl Is Nothing Then Exit Sub

    On Error Resume Next
    listCtrl.List(entryIdx - 1, 2) = _
        ShowHideStatusText(showHideEntries.IsHidden(entryIdx))
    On Error GoTo 0
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
        WarningOnSheet "ClickShowHide", "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    InitializeTrads

    If Not OpenShowHideFor(sh) Then Exit Sub

    'Read the choices of the last session and put the sheet in step with them.
    'With nothing saved yet, the sheet is the record: read it into the entries, so
    'a column the user hid by hand shows as hidden in the form.
    If LoadShowHideState(showHideEntries, activeLayout) > 0 Then
        showHideEntries.Apply activeLayout
    Else
        showHideEntries.Adopt activeLayout
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
    SaveShowHideState showHideEntries, activeLayout
    LogShowHideLine "ClickShowHide", "showhide", activeLayout, _
                    EntryCountText(showHideEntries, sh.Name)
    Set activeShowHideForm = Nothing

    ProtectAfterShowHide sh
End Sub

'Put the sheet back under protection once a show/hide session ends.
'
'ShowHideLayout brackets each write and closes the bracket itself, so the sheet
'is protected on every path the layout controls. It does not control every path:
'a raise inside the form, or a button on the form that reaches the sheet another
'way, leaves the bracket open and the user carries on typing over locked cells.
'One protect on the way out closes it whatever happened.
Private Sub ProtectAfterShowHide(ByVal sh As Worksheet)
    If sh Is Nothing Then Exit Sub
    If pass Is Nothing Then Exit Sub

    On Error Resume Next
    pass.Protect sh.Name
    On Error GoTo 0
End Sub

'Build the section context of one worksheet: the sections, the entry list and
'the layout, reconciled with what the sheet shows today. Answers Nothing on a
'sheet that carries no section map.
'
'The entry list and the layout are the module ones, so a section change made
'while the show/hide form is open lands on the same pair that form is working
'from.
Private Function SectionContextFor(ByVal sh As Worksheet, _
                                   ByVal shType As String) As SectionShowHide
    Dim layer As Byte
    Dim secMap As SectionMap

    layer = ResolveShowHideLayer(shType)
    If layer = 0 Then Exit Function

    Set secMap = SectionMap.Create(sh)
    If secMap.Count = 0 Then Exit Function

    Set showHideEntries = EntriesFor(sh, layer, DictionaryObject())
    Set activeLayout = LayoutFor(sh, layer)

    'The same reconciliation ClickShowHide does. With nothing saved yet the
    'sheet is the record, so a column the user hid by hand is read as hidden
    'and the toggle below agrees with what the user can see.
    If LoadShowHideState(showHideEntries, activeLayout) > 0 Then
        showHideEntries.Apply activeLayout
    Else
        showHideEntries.Adopt activeLayout
    End If

    Set SectionContextFor = SectionShowHide.Create(secMap, showHideEntries, _
                                                   activeLayout)
End Function

'Put the cursor somewhere that names no section, so the press that follows a
'hide is read as "bring the last one back". The first cell of the sheet is the
'go-to-section dropdown and is never on a title line.
Private Sub ParkTheCursor(ByVal sh As Worksheet)
    On Error Resume Next
    sh.Cells(1, 1).Select
    On Error GoTo 0
End Sub

'@Description("Callback for click on show/hide section in a linelist worksheet")
'@EntryPoint
Public Sub ClickShowHideSection()
    Attribute ClickShowHideSection.VB_Description = "Callback for click on show/hide section in a linelist worksheet"

    Dim sh As Worksheet
    Dim shType As String
    Dim sections As SectionShowHide
    Dim sectionIdx As Long
    Dim hideIt As Boolean
    Dim titleRng As Range

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Sections are laid out on the data entry sheets alone. The printed
    'companion carries the same columns and the CRF its own rows, and
    'SectionBuilder writes a map for neither.
    If (shType <> "HList") And (shType <> "VList") Then
        WarningOnSheet "ClickShowHideSection", "MSG_DataSheet"
        Exit Sub
    End If

    InitializeTrads

    Set sections = SectionContextFor(sh, shType)
    If sections Is Nothing Then
        WarningOnSheet "ClickShowHideSection", "MSG_SectionTitleCell"
        Exit Sub
    End If

    'A section is hidden only when the user is standing on its title, which is
    'row 5 of a data entry sheet and column 2 of a vertical one. Anywhere else
    'the press brings back the section the last press collapsed.
    sectionIdx = sections.SectionAtCell(SelectedCells())

    If sectionIdx > 0 Then
        hideIt = True
    ElseIf StrComp(lastHiddenSectionSheet, sh.Name, vbTextCompare) = 0 Then
        sectionIdx = lastHiddenSectionIndex
        hideIt = False
    End If

    If sectionIdx = 0 Then
        WarningOnSheet "ClickShowHideSection", "MSG_SectionTitleCell"
        GoTo CleanUp
    End If

    If Not sections.CanChange(sectionIdx) Then
        WarningOnSheet "ClickShowHideSection", "MSG_SectionTitleCell"
        GoTo CleanUp
    End If

    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState

    sections.SetHidden sectionIdx, hideIt
    SaveShowHideState showHideEntries, activeLayout

    If hideIt Then
        lastHiddenSectionSheet = sh.Name
        lastHiddenSectionIndex = sectionIdx
        ParkTheCursor sh
    Else
        lastHiddenSectionSheet = vbNullString
        lastHiddenSectionIndex = 0
        'Land on the title of the section that came back, so the next press
        'collapses it again and the user sees where it went.
        Set titleRng = sections.TitleCell(sectionIdx)
        If Not titleRng Is Nothing Then titleRng.Select
    End If

    'Named while the section context is still standing: the CleanUp label below
    'drops it, and a log line reading "showhide-section: sheet1" is the same
    'line whichever of a dozen sections the user just collapsed.
    LogShowHideLine "ClickShowHideSection", "showhide-section", activeLayout, _
                    SectionMoveText(sections, sectionIdx, hideIt, sh.Name)

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "ClickShowHideSection", "showhide-section", _
                                           Err.Description
    LinelistEventsManager.LLExitBusyState
    ProtectAfterShowHide sh

CleanUp:
    'The show/hide form works from this pair while it is open, so it is left
    'alone when the section change came from the sections form.
    If activeShowHideForm Is Nothing Then
        Set showHideEntries = Nothing
        Set activeLayout = Nothing
    Else
        PopulateShowHideList activeShowHideForm
    End If
End Sub

'What the user has selected, or Nothing when the selection is not a range. A
'chart or a shape holds the selection as an object of its own.
Private Function SelectedCells() As Range
    If TypeName(Application.Selection) <> "Range" Then Exit Function
    Set SelectedCells = Application.Selection
End Function


'@section The sections form
'===============================================================================
'The section button acts on one section, the one the cursor stands on.
'F_ShowHideSections offers the same two actions over the whole list, so a user
'who wants to collapse a section far from the cursor never has to travel to it.
'The form is opened from the show/hide form and works on the same entry list and
'the same layout.

'The title one section shows in the list.
'
'A run of variables the dictionary left with no `main section` is recorded as a
'block like any other, so it reaches this list with an empty title. Two thirds
'of one data entry sheet can be made of them, and a blank first column read as
'though the status column had grown rows of its own. They are real blocks and
'they stay hideable, so they are listed under a translated stand-in instead.
'
'The tag itself is the fallback, which is what TranslatedValue answers for a
'code the translation table has no row for. A linelist generated before the row
'reached the workbook reads "LBL_EmptySection" and still works.
Private Function SectionDisplayName(ByVal sectionName As String) As String
    SectionDisplayName = sectionName
    If LenB(Trim$(sectionName)) > 0 Then Exit Function

    SectionDisplayName = EMPTY_SECTION_TAG
    If tradsform Is Nothing Then Exit Function

    On Error Resume Next
    SectionDisplayName = tradsform.TranslatedValue(EMPTY_SECTION_TAG)
    On Error GoTo 0
End Function

'Fill the sections list. Two columns: the title, and whether the section is
'shown or hidden right now. A section holding nothing the user owns is listed
'and its status is left empty, so a reader can see it is there and that the
'form will not move it.
Private Sub PopulateSectionsList(ByVal frm As Object)
    Dim listCtrl As Object
    Dim counter As Long
    Dim rowIdx As Long

    If frm Is Nothing Then Exit Sub
    If activeSections Is Nothing Then Exit Sub

    On Error Resume Next
    Set listCtrl = frm.Controls("LST_Sections")
    On Error GoTo 0

    If listCtrl Is Nothing Then Exit Sub

    On Error Resume Next
    listCtrl.ColumnCount = 2
    On Error GoTo 0

    listCtrl.Clear
    For counter = 1 To activeSections.Count
        rowIdx = counter - 1
        listCtrl.AddItem SectionDisplayName(activeSections.SectionNameAt(counter))

        On Error Resume Next
        If activeSections.CanChange(counter) Then
            listCtrl.List(rowIdx, 1) = _
                ShowHideStatusText(activeSections.IsHidden(counter))
        End If
        On Error GoTo 0
    Next
End Sub

'@Description("Open the sections form from the show/hide form")
'@EntryPoint
Public Sub ClickOpenShowHideSections()
    Attribute ClickOpenShowHideSections.VB_Description = "Open the sections form from the show/hide form"

    Dim sh As Worksheet
    Dim shType As String

    On Error GoTo ErrHand

    'The show/hide form is open on top of its own sheet, and that is the sheet
    'the sections belong to. ActiveSheet answers the same thing while the form
    'is up, and it is what a press from the ribbon has.
    If activeLayout Is Nothing Then
        Set sh = ActiveSheet
    Else
        Set sh = activeLayout.Wksh
    End If

    shType = SheetTag(sh)

    InitializeTrads

    'Sections are laid out on the data entry sheets alone.
    If (shType <> "HList") And (shType <> "VList") Then
        WarningOnSheet "ClickOpenShowHideSections", "MSG_DataSheet"
        Exit Sub
    End If

    Set activeSections = SectionContextFor(sh, shType)
    If activeSections Is Nothing Then
        WarningOnSheet "ClickOpenShowHideSections", "MSG_SectionTitleCell"
        Exit Sub
    End If

    'The form is reached by name rather than written into the code.
    '
    'Naming F_ShowHideSections here would be the house pattern, and it is what
    'every other form call in this module does. It is also a COMPILE time
    'reference: a workbook that does not carry the form stops compiling
    'altogether, and every button of the linelist goes down with it. This module
    'travels into each generated linelist through CodeTransfer, so it would take
    'any linelist built before the form exists with it. UserForms.Add resolves
    'the name when the button is pressed, and a workbook without the form says
    'so and keeps working.
    On Error Resume Next
    Set activeSectionsForm = UserForms.Add("F_ShowHideSections")
    On Error GoTo ErrHand

    If activeSectionsForm Is Nothing Then
        LogWarningLine "ClickOpenShowHideSections", "showhide-sections", "no form named F_ShowHideSections"
        WarningOnSheet "ClickOpenShowHideSections", "MSG_NoSectionsForm"
        Exit Sub
    End If

    PopulateSectionsList activeSectionsForm
    activeSectionsForm.Show

    'Every option click has already written to the sheet, so the save at the
    'foot records the state the user is looking at.
    SaveShowHideState showHideEntries, activeLayout
    'The form moves any number of sections in one session, so the line says how
    'much of the sheet is left standing rather than naming one of them.
    LogShowHideLine "ClickOpenShowHideSections", "showhide-sections", activeLayout, _
                    SectionCountText(activeSections, sh.Name)
    ProtectAfterShowHide sh

    'The show/hide form underneath lists the variables one by one, and a whole
    'section just moved.
    If Not activeShowHideForm Is Nothing Then PopulateShowHideList activeShowHideForm

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "ClickOpenShowHideSections", "showhide-sections", Err.Description

    'UserForms.Add builds a fresh instance and Hide keeps it alive, so the
    'instance is unloaded here. Two opens would otherwise leave two of them
    'standing for as long as the workbook is open.
    On Error Resume Next
    Unload activeSectionsForm
    On Error GoTo 0

    Set activeSections = Nothing
    Set activeSectionsForm = Nothing
End Sub

'@Description("Callback for click on the list of the sections form")
'@EntryPoint
Public Sub ClickListShowHideSections(ByVal Index As Long)
    Attribute ClickListShowHideSections.VB_Description = "Callback for click on the list of the sections form"

    Dim sectionIdx As Long
    Dim canChange As Boolean

    If activeSections Is Nothing Then Exit Sub
    If activeSectionsForm Is Nothing Then Exit Sub

    sectionIdx = Index + 1
    If sectionIdx < 1 Or sectionIdx > activeSections.Count Then Exit Sub

    canChange = activeSections.CanChange(sectionIdx)

    On Error Resume Next
    If activeSections.IsHidden(sectionIdx) Then
        activeSectionsForm.OPT_Hide.Value = True
    Else
        activeSectionsForm.OPT_Show.Value = True
    End If
    activeSectionsForm.OPT_Show.Enabled = canChange
    activeSectionsForm.OPT_Hide.Enabled = canChange
    On Error GoTo 0
End Sub

'@Description("Callback for click on the show and hide options of the sections form")
'@EntryPoint
Public Sub ClickOptionsShowHideSections(ByVal Index As Long)
    Attribute ClickOptionsShowHideSections.VB_Description = "Callback for click on the show and hide options of the sections form"

    Dim sectionIdx As Long
    Dim shouldHide As Boolean
    Dim listCtrl As Object

    If activeSections Is Nothing Then Exit Sub
    If activeSectionsForm Is Nothing Then Exit Sub

    sectionIdx = Index + 1
    If sectionIdx < 1 Or sectionIdx > activeSections.Count Then Exit Sub
    If Not activeSections.CanChange(sectionIdx) Then Exit Sub

    shouldHide = activeSectionsForm.OPT_Hide.Value

    On Error GoTo ErrHand
    LinelistEventsManager.LLEnterBusyState

    activeSections.SetHidden sectionIdx, shouldHide

    'The section button reads this pair, so a section collapsed from the form
    'is the one a press of the button brings back.
    If shouldHide Then
        lastHiddenSectionSheet = activeSections.Wksh.Name
        lastHiddenSectionIndex = sectionIdx
    ElseIf lastHiddenSectionIndex = sectionIdx Then
        lastHiddenSectionSheet = vbNullString
        lastHiddenSectionIndex = 0
    End If

    On Error Resume Next
    Set listCtrl = activeSectionsForm.Controls("LST_Sections")
    If Not listCtrl Is Nothing Then
        listCtrl.List(sectionIdx - 1, 1) = _
            ShowHideStatusText(activeSections.IsHidden(sectionIdx))
    End If
    On Error GoTo 0

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "ClickOptionsShowHideSections", "showhide-sections", Err.Description
    LinelistEventsManager.LLExitBusyState
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
        If Not activeLayout Is Nothing Then
            isVertical = activeLayout.IsVertical(showHideEntries.PositionIndex(entryIdx))
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
    If activeLayout Is Nothing Then Exit Sub
    If activeShowHideForm Is Nothing Then Exit Sub

    entryIdx = Index + 1
    If entryIdx < 1 Or entryIdx > showHideEntries.EntryCount Then Exit Sub
    If Not showHideEntries.IsFree(entryIdx) Then Exit Sub

    shouldHide = activeShowHideForm.OPT_Hide.Value
    showHideEntries.SetHidden entryIdx, shouldHide

    posIdx = showHideEntries.PositionIndex(entryIdx)
    If posIdx = 0 Then Exit Sub

    activeLayout.BeginBatch
    activeLayout.SetHidden posIdx, shouldHide

    'The header direction belongs to the two show buttons of the printed form.
    'Choosing Hide leaves the direction as it was, so showing the column again
    'brings back the header the user had.
    If activeLayout.SupportsOrientation And Not shouldHide Then
        wantsVertical = False
        On Error Resume Next
        wantsVertical = activeShowHideForm.OPT_PrintShowVerti.Value
        On Error GoTo 0
        activeLayout.SetOrientation posIdx, wantsVertical
    End If

    activeLayout.EndBatch

    'The status column of the row the user just changed
    RefreshShowHideRow entryIdx
End Sub

'@Description("Callback for click on column width in show/hide")
'@EntryPoint
Public Sub ClickColWidth(ByVal Index As Long)
    Attribute ClickColWidth.VB_Description = "Callback for click on column width in show/hide"

    Dim entryIdx As Long
    Dim posIdx As Long
    Dim inputValue As String

    If showHideEntries Is Nothing Then Exit Sub
    If activeLayout Is Nothing Then Exit Sub

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

    activeLayout.SetSize posIdx, CDbl(inputValue)
End Sub


'@Description("Callback for click on the Print Button")
'@EntryPoint
Public Sub ClickOpenPrint()
    Attribute ClickOpenPrint.VB_Description = "Callback for click on the Print Button"

    Dim sh As Worksheet
    Dim printsh As Worksheet
    Dim shType As String
    Dim openedName As String

    On Error GoTo ErrOpen

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    InitializeTrads

    If shType <> "HList" Then
        WarningOnSheet "ClickOpenPrint", "MSG_DataSheet"
        Exit Sub
    End If

    Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)

    'Held as text. The log line below is reached on the error path too, where
    'the sheet may never have been resolved, and reading .Name off Nothing
    'there would raise 91 inside the handler.
    openedName = printsh.Name

    'UnProtect current workbook
    pass.UnProtect wb
    'Unhide the linelist Print
    printsh.Visible = xlSheetVisible
    printsh.Activate

ErrOpen:
    LogOutcomeLine "ClickOpenPrint", "open-print", Err.Number, Err.Description, _
                   openedName
    pass.Protect wb
End Sub


'@Description("Callback for click on the CRF Button")
'@EntryPoint
Public Sub ClickOpenCRF()
    Attribute ClickOpenCRF.VB_Description = "Callback for click on the CRF Button"

    Dim sh As Worksheet
    Dim crfsh As Worksheet
    Dim shType As String
    Dim openedName As String

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    InitializeTrads

    If shType <> "HList" Then
        WarningOnSheet "ClickOpenCRF", "MSG_DataSheet"
        Exit Sub
    End If

    'The CRF companion may simply not have been built. LLDataEntry.Build makes
    'the print_ companion and nothing else -- no AddOutputSheet call in this
    'project passes sheetScope:=3 -- so on a linelist generated today this sheet
    'is always absent. It used to be looked up inside the handler below, which
    'turned the missing sheet into error 9, wrote a log line nobody reads and
    'returned as though the button had worked. Saying so is the whole fix here;
    'building the sheet is a separate piece of work.
    On Error Resume Next
        Set crfsh = wb.Worksheets(CRFPREFIX & sh.Name)
    On Error GoTo 0

    'MSG_NoCRFSheet, because the user IS standing on a data entry sheet and
    'MSG_DataSheet told them to go and find one. The sheet is what is missing.
    If crfsh Is Nothing Then
        LogWarningLine "ClickOpenCRF", "open-crf", "no worksheet named " & CRFPREFIX & sh.Name
        WarningOnSheet "ClickOpenCRF", "MSG_NoCRFSheet"
        Exit Sub
    End If

    openedName = crfsh.Name

    On Error GoTo ErrOpen

    'UnProtect current workbook
    pass.UnProtect wb
    'Unhide the CRF companion
    crfsh.Visible = xlSheetVisible
    crfsh.Activate

ErrOpen:
    LogOutcomeLine "ClickOpenCRF", "open-crf", Err.Number, Err.Description, _
                   openedName
    pass.Protect wb
End Sub

'@Description("Callback for click on close worksheet: print, CRF, log or import report")
'@EntryPoint
Public Sub ClickCloseSheet()
    Attribute ClickCloseSheet.VB_Description = "Callback for click on close worksheet"

    Dim sh As Worksheet
    Dim shType As String
    Dim printsh As Worksheet
    Dim crfsh As Worksheet
    Dim logSheetName As String
    Dim reportSheetName As String
    Dim actionCode As String
    Dim closedName As String

    On Error GoTo ErrClose
    actionCode = "close-print"
    Set sh = ActiveSheet
    closedName = sh.Name

    InitializeTrads

    shType = SheetTag(sh)

    'Neither the log nor the import report carries a sheet tag; the name is the
    'mark for both. An import leaves its report on screen and the user is given
    'no way to put it away, so this button knows it too.
    logSheetName = LLLog.SheetName
    reportSheetName = ImportChecking.ReportSheetName()

    If shType <> "HList" And shType <> "HList Print" And shType <> "HList CRF" _
       And sh.Name <> logSheetName And sh.Name <> reportSheetName Then
        WarningOnSheet "ClickCloseSheet", "MSG_PrintCRFOrDataSheet"
        Exit Sub
    End If

    'Unprotect workbook
    pass.UnProtect wb

    If sh.Name = logSheetName Then
        actionCode = "close-log"
        sh.Visible = xlSheetVeryHidden
    ElseIf sh.Name = reportSheetName Then
        'Very hidden, like the log. ShowImportCheckings puts it back on screen
        'on the next import, so nothing is lost by keeping it out of the tab bar
        'and out of the unhide list.
        actionCode = "close-import-report"
        sh.Visible = xlSheetVeryHidden
    ElseIf shType = "HList" Then
        'A linelist generated before the CRF companion was built carries the
        'print sheet alone, and looking the missing one up used to raise 9 and
        'take the print sheet's close down with it.
        On Error Resume Next
        Set printsh = wb.Worksheets(PRINTPREFIX & sh.Name)
        Set crfsh = wb.Worksheets(CRFPREFIX & sh.Name)
        On Error GoTo ErrClose

        If Not printsh Is Nothing Then printsh.Visible = xlSheetVeryHidden
        If Not crfsh Is Nothing Then crfsh.Visible = xlSheetVeryHidden
    ElseIf shType = "HList Print" Then
        Set printsh = sh
        printsh.Visible = xlSheetVeryHidden
    ElseIf shType = "HList CRF" Then
        Set crfsh = sh
        crfsh.Visible = xlSheetVeryHidden
    End If


ErrClose:
    'A press on a data entry sheet puts its print and CRF companions away, so
    'the line names the sheet the press was made on rather than what went
    'hidden: that one name is what the reader can act on.
    LogOutcomeLine "ClickCloseSheet", actionCode, Err.Number, Err.Description, _
                   closedName
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
    Dim openedSheetName As String

    Set sh = ActiveSheet

    shType = SheetTag(sh)

    InitializeTrads

    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "ClickRotateAll", "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    If shType = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)

    On Error GoTo ErrHand

    'Unprotect the sheet if it is protected. The name is held, because sh is
    'the print companion on a press made from the data entry sheet and the
    'protect at the foot has to close the sheet that was opened.
    openedSheetName = sh.Name
    pass.UnProtect openedSheetName
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
    LogOutcomeLine "ClickRotateAll", "rotate", Err.Number, Err.Description, _
                   openedSheetName
    ReprotectOpenedSheet openedSheetName
    LinelistEventsManager.LLExitBusyState
End Sub

'Put back the protection a button took off one sheet.
'
'Three buttons write to a sheet property that protection guards -- the header
'direction, the row height and the column widths -- and all three left the sheet
'open when they were done. On a press made from a data entry sheet the sheet
'they open is the print companion, so the name is carried rather than read back
'off ActiveSheet.
Private Sub ReprotectOpenedSheet(ByVal sheetName As String)
    If LenB(sheetName) = 0 Then Exit Sub
    If pass Is Nothing Then Exit Sub

    On Error Resume Next
    pass.Protect sheetName
    On Error GoTo 0
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
    Dim openedSheetName As String

    Set sh = ActiveSheet

    shType = SheetTag(sh)

    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "ClickRowHeight", "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    On Error GoTo ErrHand

    InitializeTrads
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow

    If shType = "HList" Then Set sh = wb.Worksheets(PRINTPREFIX & sh.Name)
    openedSheetName = sh.Name
    pass.UnProtect openedSheetName

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
    LogSuccessLine "ClickRowHeight", "row-height", _
                   inputValue & " on " & openedSheetName

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "ClickRowHeight", "row-height", Err.Description
    ReprotectOpenedSheet openedSheetName
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
        WarningOnSheet "ClickRemoveFilters", "MSG_PrintOrDataSheet"
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
        LogSuccessLine "ClickRemoveFilters", "remove-filters", sh.Name
    End If
    Exit Sub

ErrHand:
    LogFailureLine "ClickRemoveFilters", "remove-filters", Err.Description
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
    Dim failureDetail As String

    On Error GoTo errAddRows
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow
    InitializeTrads

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet. The busy state is already on,
    'so the refusal leaves through the cleanup label below.
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "ClickAddRows", "MSG_PrintOrDataSheet"
        GoTo Cleanup
    End If

    'The busy state keeps events off while the rows come in.
    pass.UnProtect "_active"

    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)
    nbRows = IIf(shType = "HList", 199, 10)
    csTab.AddRows nbRows:=nbRows
    LogSuccessLine "ClickAddRows", "add-rows", nbRows & " rows on " & sh.Name

Cleanup:
    'Whichever of the two sheets was opened is closed again. The test used to
    'read `shType = "HList"`, so rows added on a print sheet unprotected it and
    'walked away.
    If shType = "HList" Or shType = "HList Print" Then pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    Exit Sub

errAddRows:
    'The reason is copied first. The two cleanup calls below carry On Error
    'statements of their own, and those clear Err, so a reason read after them
    'is always empty and the log line said only that the button failed.
    failureDetail = Err.Description

    On Error Resume Next
    If shType = "HList" Or shType = "HList Print" Then pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    FailureOnSheet "ClickAddRows", "MSG_ErrAddRows", logDetail:=failureDetail
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
    Dim failureDetail As String

    On Error GoTo errDelRows
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlWait
    InitializeTrads

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet. The busy state is already on,
    'so the refusal leaves through the cleanup label below.
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "ClickResize", "MSG_PrintOrDataSheet"
        GoTo Cleanup
    End If

    'The busy state keeps events off while the rows go out.
    pass.UnProtect "_active"

    nbBlank = BlankRowCountOf(sh)
    Set Lo = sh.ListObjects(1)
    Set csTab = CustomTable.Create(Lo)

    csTab.RemoveRows totalCount:=nbBlank
    LogSuccessLine "ClickResize", "resize", _
                   nbBlank & " blank rows kept on " & sh.Name

Cleanup:
    'Protected whatever the sheet turned out to be. The test used to read
    '`shType = "HList"`, so a resize made on a print sheet unprotected it and
    'walked away; it then read both tags, which still left the protection
    'riding on a variable read earlier in the walk. Protect is harmless on a
    'sheet that was never unprotected, and leaving one open is not.
    On Error Resume Next
    pass.Protect "_active"
    On Error GoTo 0
    LinelistEventsManager.LLExitBusyState
    Exit Sub

errDelRows:
    'Copied before the cleanup calls below, whose own On Error statements
    'clear Err.
    failureDetail = Err.Description

    On Error Resume Next
    pass.Protect "_active"
    LinelistEventsManager.LLExitBusyState
    FailureOnSheet "ClickResize", "MSG_ErrDelRows", logDetail:=failureDetail
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
    Dim failureDetail As String

    'initialize translations
    InitializeTrads

    Set expsh = wb.Worksheets(EXPORTSHEET)
    Set expObj = LLExport.Create(expsh)

    On Error GoTo errLoadExp

    'The form owns its export buttons: it adds one per active export, it holds
    'the object that answers each click, and it takes them off again below.
    'The qualified call runs the F_Export code-behind, whose source is the
    'FormLogicExport module. It answers the free position under the last
    'button, which the fixed controls are laid out from.
    'The workbook is the module-level wb, which InitializeTrads set above and
    'every other entry point of this module uses.
    topPosition = F_Export.SetupExportForm(wb, tradsmess, expObj, _
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
    'This button opens the form and takes it down again; the exports themselves
    'are logged by the form. The line therefore records the session, and names
    'the sheet the export list was read from.
    LogSuccessLine "ClickExport", "export", "export form closed on " & expsh.Name

    Exit Sub

errLoadExp:
    'Copied before the teardown below, whose own On Error statement clears Err.
    failureDetail = Err.Description

    On Error Resume Next
    F_Export.TeardownExportForm
    FailureOnSheet "ClickExport", "MSG_ErrLoadExport", logDetail:=failureDetail
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
        WarningOnSheet "ClickGeoApp", "MSG_DataOrSpatioSheet"
        Exit Sub
    End If

    InitializeTrads
    
    Select Case shType

    Case "HList"

        tabName = TableNameOf(sh)
        startRow = sh.Range(tabName & "_START").Row + 1
        targetColumn = ActiveCell.Column

        If ActiveCell.Row >= startRow Then

            hfOrGeo = ColumnControl(sh, tabName, targetColumn)
            Select Case hfOrGeo
            'geo1 ONLY, and that is not a narrow reading of the button. The form
            'writes its answer back by splitting the picked place on a separator
            'and pouring the four parts out from the cell it was opened on:
            'FormLogicGeo clears Range(cellRng, cellRng.Offset(, 3)) first. Open
            'it on adm2 and that clears adm2 through adm5 -- the fourth being
            'pcode_adm1 -- and writes all four admin names one column to the
            'right. The anchor has to be adm1, so the other three levels refuse.
            Case "geo1"
                LoadGeo GeoScopeAdmin
            Case "hf"
                LoadGeo GeoScopeHF
            Case Else
                WarningOnSheet "ClickGeoApp", "MSG_WrongCells"
            End Select
        Else
            WarningOnSheet "ClickGeoApp", "MSG_WrongCells"
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
    Dim doneCount As Long

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
        doneCount = doneCount + 1
    Next

ErrHand:
    'The count is what a failure line needs read beside it: a raise on the
    'third of the four sheets leaves the first two calculated, and the number
    'says which state the workbook is in.
    LogOutcomeLine "ClickCalculate", "calculate", Err.Number, Err.Description, _
                   doneCount & " analysis sheets calculated"
    LinelistEventsManager.LLExitBusyState
End Sub


'@Description("Print the current linelist")
'@EntryPoint
Public Sub ClickPrintLL()
    Attribute ClickPrintLL.VB_Description = "Print the current linelist"

    Dim sh As Worksheet
    Dim shType As String
    Dim failureDetail As String

    'Set up the sheet with some print Characteristics
    Set sh = ActiveSheet

    'Test to be sure we are on print or linelist worksheet
    shType = SheetTag(sh)

    'Warning if not on print or hlist worksheet
    If shType <> "HList Print" And shType <> "HList" Then
        WarningOnSheet "ClickPrintLL", "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    InitializeTrads

    'On HListSheet, open the print sheet
    If shType = "HList" Then ClickOpenPrint

    Set sh = ActiveSheet

    'Everything from here on is guarded. It used to run under On Error Resume
    'Next as far as the End With, and the PrintPreview at the bottom sat outside
    'any handler at all -- so a preview that failed raised VBA's own dialog on a
    'workbook whose user has no VBE, with PrintCommunication left False and the
    'sheet left unprotected. A print sheet with no ListObject took the same path
    'through the PrintArea line.
    On Error GoTo ErrPrint

    pass.UnProtect sh.Name

    'PrintCommunication belongs to Excel on Windows. Mac Excel carries the name
    'and refuses the write, and this line was the first statement under the
    'handler, so the whole button ended in the failure box on every Mac before
    'one page-setup property had been read. It is a speed setting: the writes
    'below all work without it, they are just sent to the driver one at a time.
    On Error Resume Next
        Application.PrintCommunication = False
    On Error GoTo ErrPrint

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
        .CenterHorizontally = True
        .CenterVertically = False
        'Landscape
        .Orientation = xlLandscape
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

        'The printer answers these two, and a printer that has no A3 tray or no
        'say on the dots per inch refuses them. Neither is worth losing the page
        'setup over, so each is asked for on its own and the sheet keeps the
        'driver's own value when the answer is no.
        On Error Resume Next
            .PaperSize = xlPaperA3
            .PrintQuality = 600
        On Error GoTo ErrPrint
    End With

    On Error Resume Next
        Application.PrintCommunication = True
    On Error GoTo ErrPrint

    sh.PrintPreview

    LogSuccessLine "ClickPrintLL", "print-preview", sh.Name
    pass.Protect sh.Name
    Exit Sub

ErrPrint:
    'The reason is copied first. Every On Error statement below clears Err, and
    'the restore needs one of its own.
    failureDetail = Err.Description

    'PrintCommunication is an APPLICATION setting. Left False it silently drops
    'every later page-setup write in this Excel session, including ones made by
    'the user by hand, so it is restored before anything else is attempted.
    On Error Resume Next
        Application.PrintCommunication = True
    On Error GoTo 0

    LogFailureLine "ClickPrintLL", "print-preview", failureDetail
    FailureOnSheet "ClickPrintLL", "MSG_PrintFailed", failureDetail
    On Error Resume Next
        pass.Protect sh.Name
    On Error GoTo 0
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

    InitializeTrads
    On Error GoTo ErrHand

    'The form fills its own list in UserForm_Activate, from the same event
    'service this module reads. The rows used to be built here and pushed into
    'the control, and before that staged on the __temp worksheet.
    [F_ShowVarLabels].Show
    Exit Sub

ErrHand:
    LogFailureLine "ClickOpenVarLab", "open-varlab", Err.Description
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
        WarningOnSheet "ClickSortTable", "MSG_DataSheet"
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
        LogSuccessLine "ClickSortTable", "sort", _
                       headerName & " " & _
                       IIf(sortOrder = xlAscending, "ascending", "descending") & _
                       " on " & sh.Name
    End If
    Exit Sub

ErrHand:
    LogFailureLine "ClickSortTable", "sort", Err.Description
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
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    InitializeTrads

    'Events stay off across the whole import. The resize below writes to every
    'HList sheet and the import writes thousands of rows after it, and each of
    'those writes would otherwise raise the sheet-change handler.
    '
    'It goes through the events manager, which is the one owner of that flag in
    'a running linelist. This sub used to write Application.EnableEvents itself,
    'with no handler over it, so a raise anywhere inside -- an unprotect
    'refused, a table that would not resize, the import itself -- ended the sub
    'with events still off, and Excel holds them off for the rest of the
    'session. A linelist in that state answers no worksheet event at all: no geo
    'cascade, no checking, no autofill. The user reads it as the dropdowns
    'having stopped working.
    '
    'It also poisoned the scope underneath it. F_Advanced.HandleImportData opens
    'an ApplicationState of its own, and that snapshot was taken AFTER the raw
    'line below had already turned events off, so the form's own restore put
    '"off" back and the only line that undid it was the last one of this sub.
    On Error GoTo ImportFailed
    LinelistEventsManager.LLEnterQuietState

    'Resize all HList tables before import (remove blank rows)
    For Each sh In wb.Worksheets
        If SheetTag(sh) = "HList" Then
            nbBlank = BlankRowCountOf(sh)
            Set Lo = sh.ListObjects(1)
            Set csTab = CustomTable.Create(Lo)
            pass.UnProtect sh.Name
            On Error Resume Next
            If Not (Lo.AutoFilter Is Nothing) Then Lo.AutoFilter.ShowAllData
            csTab.RemoveRows totalCount:=nbBlank
            On Error GoTo ImportFailed
            pass.Protect sh.Name
        End If
    Next

    'Import data using LLImporter API (handles file picker, the question about
    'the rows already entered, busy state and report). Qualified through the
    'form: the walk lives in the F_Advanced code-behind.
    F_Advanced.HandleImportData wb, tradsmess

    'Update all the listAuto in the workbook
    LinelistEventsManager.UpdateAllListAuto

ImportCleanup:
    'Opens with a suppression because a Resume leaves the handler armed.
    On Error Resume Next
    LinelistEventsManager.LLExitQuietState
    On Error GoTo 0

    'The raise is carried out rather than swallowed, so an import that failed
    'says so exactly as it did before. What changed is that the events are back
    'on by the time it does.
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Sub

ImportFailed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    Resume ImportCleanup
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
    LogOutcomeLine "ClickResetColumns", "reset-columns", Err.Number, Err.Description
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
    showHideEntries.Apply activeLayout

    'Toggle the flag
    If showOptional Then
        wkbNames.SetValue RNGSHOWALLOPTIONALS, "no"
    Else
        wkbNames.SetValue RNGSHOWALLOPTIONALS, "yes"
    End If

    SaveShowHideState showHideEntries, activeLayout

    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If

    'The success line sits above the label so a declined confirm logs nothing.
    'showOptional is what the flag said BEFORE the toggle, so the line reads
    'the way the button just went.
    LogShowHideLine "ClickShowHideMinimal", "showhide-minimal", activeLayout, _
                    "optional variables " & _
                    IIf(showOptional, "hidden", "shown") & " on " & sh.Name

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "ClickShowHideMinimal", "showhide-minimal", Err.Description
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

    If LoadShowHideState(showHideEntries, activeLayout) > 0 Then
        showHideEntries.Apply activeLayout
    Else
        showHideEntries.Adopt activeLayout
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
    Set activeLayout = printLayout
    If Not activeShowHideForm Is Nothing Then
        PopulateShowHideList activeShowHideForm
    End If

    'The success line sits above the label so a declined confirm logs nothing.
    LogShowHideLine "ClickMatchLinelistShowHide", "showhide-match", printLayout, _
                    printsh.Name & " matched to " & sheetName

ErrHand:
    If Err.Number <> 0 Then LogFailureLine "ClickMatchLinelistShowHide", "showhide-match", Err.Description
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
    Dim openedSheetName As String

    Set sh = ActiveSheet
    shType = SheetTag(sh)

    InitializeTrads

    'The print sheet holds a ListObject of its own and the same columns, and it
    'is the sheet whose widths a user actually wants to settle before printing.
    'It used to be refused here, so autofit worked on the data sheet alone.
    If shType <> "HList" And shType <> "HList Print" Then
        WarningOnSheet "clickAutoFit", "MSG_PrintOrDataSheet"
        Exit Sub
    End If

    On Error GoTo ErrHand

    'Column widths are a protected property, so the sheet has to come out of
    'protection first. The data sheet was never protected against this, which is
    'why the missing unprotect only ever showed up on the print sheet.
    openedSheetName = sh.Name
    pass.UnProtect openedSheetName
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
    LogOutcomeLine "clickAutoFit", "autofit", Err.Number, Err.Description, sh.Name
    On Error Resume Next
    ReprotectOpenedSheet openedSheetName
    LinelistEventsManager.LLExitBusyState
End Sub
