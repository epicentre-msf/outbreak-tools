Attribute VB_Name = "GeoModule"
Attribute VB_Description = "Combined geo and spatial analysis logic for the linelist"

'@Folder("Geo")
'@ModuleDescription("Combined geo and spatial analysis logic for the linelist")
'@IgnoreModule UnrecognizedAnnotation, ImplicitActiveSheetReference, UseMeaningfulName, HungarianNotation

Option Explicit
Option Private Module

'@section Constants
'===============================================================================

Private Const DROPDOWNSHEET As String = "__dropdown_lists"
Private Const SPATIALSHEET As String = "__spatial_tables"
Private Const PASSSHEET As String = "__pass"

' How many admin levels a geobase carries.
Private Const MAX_ADMIN_LEVEL As Long = 4

'@section Module-Level State
'===============================================================================

Private drop As DropdownLists
Private pass As Passwords

'@section Initialization
'===============================================================================

' @description Initialize the dropdown lists. The object is built when it is
'              missing and kept when it is there.
Private Sub InitializeGeoElements()
    Dim wb As Workbook

    Set wb = ThisWorkbook
    If drop Is Nothing Then Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))
End Sub

' @description Initialize passwords for spatial analysis events. The old shape
'              rebuilt the manager on every Validate.
Private Sub InitializeSpatialElements()
    If pass Is Nothing Then Set pass = Passwords.Create(ThisWorkbook.Worksheets(PASSSHEET))
End Sub

'@section Error Reporting
'===============================================================================

' @description Tell the user a geo operation failed. The handler that calls
'              this has already restored the application. The message comes off
'              the shared surface, so every failure box of the linelist reads
'              the same way. The fallback is what a workbook with no usable
'              translation sheet shows, and a geo failure is exactly the state
'              where that can happen.
'              The typed local is what call-signature-scan.R reads to check the
'              member. A chained call is invisible to it, and this module carries
'              no registry row, so a chain here would be checked by nothing.
'              The caller names itself in source. VBA carries no call stack to
'              read a name from, four procedures of this module report through
'              here, and a log line that names none of them leaves a reader
'              with a reason and no idea which press produced it.
Private Sub ReportGeoError(ByVal source As String, ByVal detail As String)
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    linelistEvents.Fail "MSG_ErrGeo", detail, "The geobase could not be read", _
                        source
End Sub

' @description The one geobase manager of the workbook. EventLinelist builds it
'              once and drops it in ResetCaches, so a geobase import is followed
'              by a fresh build and the level labels below are read again. The
'              answer lives in a procedure-local at every use site: a module
'              field here would hold a manager nothing invalidates.
'              The typed local is what call-signature-scan.R reads to check the
'              member, the same reason ReportGeoError above carries one.
Private Function GeoOf() As LLGeo
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    If linelistEvents Is Nothing Then Exit Function

    Set GeoOf = linelistEvents.GeoManager()
End Function

' Whether the tab strip of each frame has been put back on its first page. The
' picker keeps its state between opens through its default instance, so the tab
' it comes up on the very first time is the tab the form was saved on, which is
' whichever one the designer happened to close. The user starts on the four
' admin lists, not on the concatenated list and not on the historic.
'
' One flag per frame, because the two pickers open independently and the first
' open of each is the one that needs it. Every later open gives the user back
' the tab they left.
Private adminPageSettled As Boolean
Private facilityPageSettled As Boolean

'@section LoadGeo — Form Display
'===============================================================================

' @description Load the F_Geo form for geo or health facility scope.
'              Initializes admin lists, historic tables, and concatenated data.
' @param hfOrGeo GeoScopeAdmin (0) for geo, GeoScopeHF (1) for health facility
'@EntryPoint
Public Sub LoadGeo(ByVal hfOrGeo As Byte)
    Dim geoList As BetterArray
    Dim historicList As BetterArray
    Dim geoObj As LLGeo

    On Error GoTo ErrLoadGeo

    InitializeGeoElements

    'The manager answers Nothing when its build failed, where LLGeo.Create used
    'to raise, so the report the user reads is asked for here.
    Set geoObj = GeoOf()
    If geoObj Is Nothing Then
        ReportGeoError "LoadGeo", "The geobase manager could not be built"
        Exit Sub
    End If

    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow

    'The form survives between opens through its default instance, so the
    'lists are emptied first, whatever the geobase holds: an emptied geobase
    'must show none of the previous session's places.
    ClearLists

    'One read of each list per open. The search boxes then scan memory on
    'every keystroke, off the same cache the form reads.
    GeoFormCache.LoadFrom ThisWorkbook

    'Every control below is reached through one block. Naming F_Geo resolves
    'the default instance and looks the control up again on each line, and the
    'open touched about twenty of them.
    With F_Geo

        Select Case hfOrGeo

        Case GeoScopeAdmin
            .LBL_Adm1.Caption = geoObj.GeoNames("adm1_name")
            .LBL_Adm2.Caption = geoObj.GeoNames("adm2_name")
            .LBL_Adm3.Caption = geoObj.GeoNames("adm3_name")
            .LBL_Adm4.Caption = geoObj.GeoNames("adm4_name")

            drop.ClearList "admin2"
            drop.ClearList "admin3"
            drop.ClearList "admin4"

            If Not geoObj.HasNoData() Then
                Set geoList = geoObj.GeoLevel(LevelAdmin1, GeoScopeAdmin)
                .LST_Adm1.List = geoList.Items
            End If

            'The concatenated tab is a search surface: its list starts empty
            'and fills from the search box at three characters. Pushing the
            'whole adm4 column here was the largest single allocation of the
            'open, and an MSForms ListBox stops outright at 65536 rows.

            Set historicList = GeoFormCache.HistoricList(GeoScopeAdmin)
            .LST_Histo.List = historicList.Items

            .FRM_Facility.Visible = False
            .FRM_Geo.Visible = True
            .LBL_Fac1.Visible = False
            .LBL_Geo1.Visible = True

            ShowFirstGeoPage .FRM_Geo, adminPageSettled

        Case GeoScopeHF
            .LBL_Adm4F.Caption = geoObj.GeoNames("hf_name")
            .LBL_Adm3F.Caption = geoObj.GeoNames("adm3_name")
            .LBL_Adm2F.Caption = geoObj.GeoNames("adm2_name")
            .LBL_Adm1F.Caption = geoObj.GeoNames("adm1_name")

            If Not geoObj.HasNoData() Then
                Set geoList = geoObj.GeoLevel(LevelAdmin1, GeoScopeHF)
                .LST_AdmF1.List = geoList.Items
            End If

            'The facility concatenated list follows the admin one: empty until
            'the search box holds three characters.

            Set historicList = GeoFormCache.HistoricList(GeoScopeHF)
            .LST_HistoF.List = historicList.Items
            .FRM_Facility.Visible = True
            .FRM_Geo.Visible = False
            .LBL_Fac1.Visible = True
            .LBL_Geo1.Visible = False

            ShowFirstGeoPage .FRM_Facility, facilityPageSettled

        Case Else
            'GeoScopeBoth exists on the enum and has no layout in the form. An
            'unknown scope used to configure nothing and show the form as the
            'previous open left it.
            LinelistEventsManager.LLExitBusyState
            ReportGeoError "LoadGeo", "Unknown geo scope " & hfOrGeo
            Exit Sub

        End Select

        .TXT_Msg.Value = vbNullString
    End With

    'Exit the busy state before the modal form comes up, so it is never
    'raised over a frozen screen. Show blocks until the form hides, so it
    'stands outside the block above: nothing holds a reference to the form
    'while the form runs.
    LinelistEventsManager.LLExitBusyState
    F_Geo.Show
    Exit Sub

ErrLoadGeo:
    LinelistEventsManager.LLExitBusyState
    ReportGeoError "LoadGeo", Err.Description
End Sub

' @description Put the tab strip of one frame on its first page, the four admin
'              lists, and only the first time that frame is opened in the
'              session. The page is chosen by position: the two frames carry a
'              tab strip each and the form gives them different names, so the
'              control is found by type rather than by name.
' @param frameControl The FRM_Geo or FRM_Facility frame of the picker
' @param alreadySettled The session flag of that frame, raised here
Private Sub ShowFirstGeoPage(ByVal frameControl As Object, _
                             ByRef alreadySettled As Boolean)
    Dim ctrl As Object

    If alreadySettled Then Exit Sub
    alreadySettled = True

    ' A frame carrying no tab strip has nothing to settle, and a tab strip that
    ' refuses the page leaves the picker on the one it was already showing.
    ' Neither is worth stopping an open for.
    On Error Resume Next
    For Each ctrl In frameControl.Controls
        If TypeName(ctrl) = "MultiPage" Then ctrl.Value = 0
    Next
    On Error GoTo 0
End Sub

' @description Empty every list control of the F_Geo form, entries and
'              selection both. Assigning a Value outside the list entries is
'              the documented route to error 380, so nothing is cleared that
'              way.
Private Sub ClearLists()
    Dim counter As Long

    With F_Geo
        .LST_Adm1.Clear
        .LST_AdmF1.Clear
        .LST_ListeAgre.Clear
        .LST_ListeAgreF.Clear
        .LST_Histo.Clear
        .LST_HistoF.Clear
        For counter = 2 To 4
            .Controls("LST_Adm" & counter).Clear
            .Controls("LST_AdmF" & counter).Clear
        Next
    End With
End Sub

'@section Admin Cascade
'===============================================================================

' @description Fill the admin list of one cascade level from the levels above
'              it, in the geo or the facility scope. The lists from the given
'              level down are emptied first, because they hold children of a
'              selection that just changed. The caption joins the parents with
'              the selected value: the geo scope reads admin1 first and the
'              facility scope reads the deepest level first, which is the
'              order CMD_Copier_Click splits back out.
' @param level Cascade level to fill, 2 to 4
' @param selectedValue The value clicked at the level above
' @param scope GeoScopeAdmin (0) or GeoScopeHF (1)
' @param separator Separator of the caption
'@EntryPoint
Public Sub ShowAdminList(ByVal level As Long, ByVal selectedValue As String, _
                         Optional ByVal scope As Byte = GeoScopeAdmin, _
                         Optional ByVal separator As String = " | ")

    Dim adminTable As BetterArray
    Dim adminNames As BetterArray
    Dim parentValues() As String
    Dim listPrefix As String
    Dim caption As String
    Dim counter As Long
    Dim levelWanted As Byte
    Dim geoObj As LLGeo

    On Error GoTo ErrShowAdmin
    Application.Cursor = xlNorthwestArrow

    'The form outlives any dead module state, so the manager is read fresh on
    'every click.
    Set geoObj = GeoOf()
    If geoObj Is Nothing Then
        Application.Cursor = xlNorthwestArrow
        ReportGeoError "ShowAdminList", "The geobase manager could not be built"
        Exit Sub
    End If

    'GeoScopeBoth exists on the enum and has no lists in the form. An
    'unknown scope used to mean facility in silence.
    If scope <> GeoScopeAdmin And scope <> GeoScopeHF Then _
        Err.Raise 5, "GeoModule", "The cascade knows no scope " & scope

    Select Case level
    Case 2
        levelWanted = LevelAdmin2
    Case 3
        levelWanted = LevelAdmin3
    Case 4
        levelWanted = LevelAdmin4
    Case Else
        Err.Raise 5, "GeoModule", "The cascade knows no level " & level
    End Select

    listPrefix = IIf(scope = GeoScopeAdmin, "LST_Adm", "LST_AdmF")

    With F_Geo
        For counter = level To MAX_ADMIN_LEVEL
            .Controls(listPrefix & counter).Clear
        Next

        'The parents above the clicked value are read off their lists. A
        'list with nothing selected answers Null, and refilling a list
        'drops its selection, so the read has to survive both.
        ReDim parentValues(1 To level - 1)
        For counter = 1 To level - 2
            parentValues(counter) = ListValueOf(.Controls(listPrefix & counter))
        Next
        parentValues(level - 1) = selectedValue

        Set adminNames = New BetterArray
        adminNames.LowerBound = 1
        For counter = 1 To level - 1
            adminNames.Push parentValues(counter)
        Next

        If scope = GeoScopeAdmin Then
            caption = parentValues(1)
            For counter = 2 To level - 1
                caption = caption & separator & parentValues(counter)
            Next
        Else
            caption = parentValues(level - 1)
            For counter = level - 2 To 1 Step -1
                caption = caption & separator & parentValues(counter)
            Next
        End If

        'Admin 2 wants the name of its one parent as a single value, and the
        'deeper levels want the table of names. GuardLevelNames holds that
        'line on the LLGeo side.
        If level = 2 Then
            Set adminTable = geoObj.GeoLevel(levelWanted, scope, selectedValue)
        Else
            Set adminTable = geoObj.GeoLevel(levelWanted, scope, adminNames)
        End If

        .TXT_Msg.Value = caption

        If adminTable.Length > 0 Then
            .Controls(listPrefix & level).List = adminTable.Items
        End If
    End With

    'The arrow is the standing cursor of a linelist session, set at open by
    'EventLinelist.OnWorkbookOpen. Leaving the default cursor here made the
    'pointer change under the geo form.
    Application.Cursor = xlNorthwestArrow
    Exit Sub

ErrShowAdmin:
    Application.Cursor = xlNorthwestArrow
    ReportGeoError "ShowAdminList", Err.Description
End Sub

' @description The value of one list control, as a string. A list with
'              nothing selected answers Null.
Private Function ListValueOf(ByVal listControl As Object) As String
    If IsNull(listControl.Value) Then Exit Function
    ListValueOf = CStr(listControl.Value)
End Function

'@section Spatial Table Updates
'===============================================================================

' @description Update all spatial tables from HList filtered data.
'              The spatial refresh button reaches this sub bare through its
'              shape's OnAction, so it holds the shared busy state and a
'              handler itself. The busy depth counts, so the wrap that
'              ClickCalculate puts around the same call nests cleanly.
'@EntryPoint
Public Sub UpdateSpTables()
    Dim sp As LLSpatial

    On Error GoTo ErrUpdate
    LinelistEventsManager.LLEnterBusyState

    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))

    UpdateFilterTables calculate:=False

    sp.Update

    LinelistEventsManager.LLExitBusyState
    Exit Sub

ErrUpdate:
    LinelistEventsManager.LLExitBusyState
    ReportGeoError "UpdateSpTables", Err.Description
End Sub

'@section Spatio-Temporal Formula Updates
'===============================================================================

' @description Update formulas in spatio-temporal tables when admin level changes.
'              Runs after the user validates a place on an SPT analysis sheet.
'              The section walk is LLSpatial.MigrateSection, so the harness
'              measures it through TestLLSpatial: every formula column of the
'              section moves from the previous admin level's concat column to
'              the new one, and a plain formula stays plain while an array one
'              stays an array one. This sub keeps the event side: the busy
'              state, the active sheet, the protection pair and the report.
' @param rngName Named range of the admin level selector
' @param actAdm New admin level (number of admin levels selected)
'@EntryPoint
Public Sub UpdateSpatioTemporalFormulas(ByVal rngName As String, _
                                        ByVal actAdm As Long)
    Dim tabId As String
    Dim prevAdm As Long
    Dim sh As Worksheet
    Dim sp As LLSpatial
    Dim unprotected As Boolean

    'The handler is armed first, so a raise in the busy-state entry or in
    'InitializeSpatialElements reaches ErrSPT and restores the application.
    On Error GoTo ErrSPT
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow
    InitializeSpatialElements

    'An unnamed active cell hands an empty rngName over, and AnalysisRanges
    'answers an empty id for any name it did not build. Both shapes used to
    'be sliced by position, which raised into the handler.
    tabId = AnalysisRanges.IdOfSpatialInput(rngName)
    If LenB(tabId) = 0 Then
        LinelistEventsManager.LLExitBusyState
        Exit Sub
    End If

    Set sh = ActiveSheet
    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))

    'The level is read and checked above the UnProtect, so a bad level
    'raises while a deliberately open sheet is still open.
    prevAdm = sp.PreviousSectionLevel(sh, rngName, tabId)

    'The caller fires on every Validate with no idea whether the level
    'changed.
    If prevAdm = actAdm Then
        LinelistEventsManager.LLExitBusyState
        Exit Sub
    End If

    pass.UnProtect "_active"
    unprotected = True

    sp.MigrateSection sh, rngName, tabId, prevAdm, actAdm

    pass.Protect sh, allowShapes:=True
    LinelistEventsManager.LLExitBusyState
    Exit Sub

ErrSPT:
    'The protection is put back only when this run took it off, so a raise
    'above the UnProtect leaves a deliberately open sheet open.
    If unprotected Then pass.Protect sh, allowShapes:=True
    LinelistEventsManager.LLExitBusyState
    ReportGeoError "UpdateSpatioTemporalFormulas", Err.Description
End Sub
