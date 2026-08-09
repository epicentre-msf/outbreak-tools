Attribute VB_Name = "GeoModule"
Attribute VB_Description = "Combined geo and spatial analysis logic for the linelist"

'@Folder("Geo")
'@ModuleDescription("Combined geo and spatial analysis logic for the linelist")
'@IgnoreModule UnrecognizedAnnotation, ImplicitActiveSheetReference, UseMeaningfulName, HungarianNotation

Option Explicit
Option Private Module

'@section Constants
'===============================================================================

Private Const GEOSHEET As String = "Geo"
Private Const DROPDOWNSHEET As String = "dropdown_lists__"
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const PASSSHEET As String = "__pass"

' How many admin levels a geobase carries.
Private Const MAX_ADMIN_LEVEL As Long = 4

'@section Module-Level State
'===============================================================================

Private geo As LLGeo
Private drop As DropdownLists
Private pass As Passwords

'@section Initialization
'===============================================================================

' @description Initialize geo elements: LLGeo and the dropdowns. Each object
'              is built when it is missing and kept when it is there.
'              LLGeo.Create walks two whole Names collections, and the old
'              shape paid that walk on every form open by overwriting the
'              module state each call.
Private Sub InitializeGeoElements()
    Dim wb As Workbook

    Set wb = ThisWorkbook
    If geo Is Nothing Then Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))
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
Private Sub ReportGeoError(ByVal detail As String)
    Dim linelistEvents As EventLinelist

    Set linelistEvents = LinelistEventsManager.EventLinelistService()
    linelistEvents.Fail "MSG_ErrGeo", detail, "The geobase could not be read"
End Sub

'@section LoadGeo — Form Display
'===============================================================================

' @description Load the F_Geo form for geo or health facility scope.
'              Initializes admin lists, historic tables, and concatenated data.
' @param hfOrGeo GeoScopeAdmin (0) for geo, GeoScopeHF (1) for health facility
'@EntryPoint
Public Sub LoadGeo(ByVal hfOrGeo As Byte)
    Dim geoList As BetterArray
    Dim historicList As BetterArray

    On Error GoTo ErrLoadGeo

    InitializeGeoElements
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow

    'The form survives between opens through its default instance, so the
    'lists are emptied first, whatever the geobase holds: an emptied geobase
    'must show none of the previous session's places.
    ClearLists

    'One read of each list per open. The search boxes then scan memory on
    'every keystroke, off the same cache the form reads.
    GeoFormCache.LoadFrom ThisWorkbook

    Select Case hfOrGeo

    Case GeoScopeAdmin
        F_Geo.LBL_Adm1.Caption = geo.GeoNames("adm1_name")
        F_Geo.LBL_Adm2.Caption = geo.GeoNames("adm2_name")
        F_Geo.LBL_Adm3.Caption = geo.GeoNames("adm3_name")
        F_Geo.LBL_Adm4.Caption = geo.GeoNames("adm4_name")

        drop.ClearList "admin2"
        drop.ClearList "admin3"
        drop.ClearList "admin4"

        If Not geo.HasNoData() Then
            Set geoList = geo.GeoLevel(LevelAdmin1, GeoScopeAdmin)
            F_Geo.LST_Adm1.List = geoList.Items
        End If

        'The concatenated tab is a search surface: its list starts empty and
        'fills from the search box at three characters. Pushing the whole
        'adm4 column here was the largest single allocation of the open, and
        'an MSForms ListBox stops outright at 65536 rows.

        Set historicList = GeoFormCache.HistoricList(GeoScopeAdmin)
        F_Geo.LST_Histo.List = historicList.Items

        F_Geo.FRM_Facility.Visible = False
        F_Geo.FRM_Geo.Visible = True
        F_Geo.LBL_Fac1.Visible = False
        F_Geo.LBL_Geo1.Visible = True

    Case GeoScopeHF
        F_Geo.LBL_Adm4F.Caption = geo.GeoNames("hf_name")
        F_Geo.LBL_Adm3F.Caption = geo.GeoNames("adm3_name")
        F_Geo.LBL_Adm2F.Caption = geo.GeoNames("adm2_name")
        F_Geo.LBL_Adm1F.Caption = geo.GeoNames("adm1_name")

        If Not geo.HasNoData() Then
            Set geoList = geo.GeoLevel(LevelAdmin1, GeoScopeHF)
            F_Geo.LST_AdmF1.List = geoList.Items
        End If

        'The facility concatenated list follows the admin one: empty until
        'the search box holds three characters.

        Set historicList = GeoFormCache.HistoricList(GeoScopeHF)
        F_Geo.LST_HistoF.List = historicList.Items
        F_Geo.FRM_Facility.Visible = True
        F_Geo.FRM_Geo.Visible = False
        F_Geo.LBL_Fac1.Visible = True
        F_Geo.LBL_Geo1.Visible = False

    Case Else
        'GeoScopeBoth exists on the enum and has no layout in the form. An
        'unknown scope used to configure nothing and show the form as the
        'previous open left it.
        LinelistEventsManager.LLExitBusyState
        ReportGeoError "Unknown geo scope " & hfOrGeo
        Exit Sub

    End Select

    'Exit the busy state before the modal form comes up, so it is never
    'raised over a frozen screen.
    LinelistEventsManager.LLExitBusyState

    F_Geo.TXT_Msg.Value = vbNullString
    F_Geo.Show
    Exit Sub

ErrLoadGeo:
    LinelistEventsManager.LLExitBusyState
    ReportGeoError Err.Description
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

    On Error GoTo ErrShowAdmin
    Application.Cursor = xlNorthwestArrow

    'Module state is lost on any unhandled error and any edit to the
    'project, and the form outlives both.
    If geo Is Nothing Then InitializeGeoElements

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
            Set adminTable = geo.GeoLevel(levelWanted, scope, selectedValue)
        Else
            Set adminTable = geo.GeoLevel(levelWanted, scope, adminNames)
        End If

        .TXT_Msg.Value = caption

        If adminTable.Length > 0 Then
            .Controls(listPrefix & level).List = adminTable.Items
        End If
    End With

    Application.Cursor = xlDefault
    Exit Sub

ErrShowAdmin:
    Application.Cursor = xlDefault
    ReportGeoError Err.Description
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
'              The busy state belongs to the caller. ClickCalculate already
'              wraps this call and restores the application on every path, and
'              UpdateFilterTables below enters the shared busy state itself.
'@EntryPoint
Public Sub UpdateSpTables()
    Dim sp As LLSpatial
    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))

    UpdateFilterTables calculate:=False

    sp.Update
End Sub

'@section Spatio-Temporal Formula Updates
'===============================================================================

' @description Update formulas in spatio-temporal tables when admin level changes.
'              Runs after the user validates a place on an SPT analysis sheet:
'              every formula column of the section moves from the previous
'              admin level's concat column to the new one. The cell loop is
'              LLSpatial.RewriteFormulas, so a plain formula stays plain and
'              an array one stays an array one.
' @param rngName Named range of the admin level selector
' @param actAdm New admin level (number of admin levels selected)
'@EntryPoint
Public Sub UpdateSpatioTemporalFormulas(ByVal rngName As String, _
                                        ByVal actAdm As Long)
    Dim tabId As String
    Dim prevRaw As Variant
    Dim prevAdm As Long
    Dim sh As Worksheet
    Dim counter As Long
    Dim headerRng As Range
    Dim valuesRng As Range
    Dim headerTexts As Variant
    Dim headerCellName As String
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
    Set headerRng = sh.Range("SPT_FORMULA_COLUMN_" & tabId)

    'The neighbour cell records the level the sheet stands on. A value
    'outside 1 to 4 means it was cleared or overwritten; rewriting on it
    'would migrate nothing while still recording the new level, and the
    'sheet would then be one level behind with no way back.
    prevRaw = sh.Range(rngName).Offset(, 1).Value
    If Not IsNumeric(prevRaw) Then _
        Err.Raise 5, "GeoModule", _
                  "The previous admin level of " & tabId & " reads " & _
                  Chr$(34) & prevRaw & Chr$(34)

    prevAdm = CLng(prevRaw)
    If prevAdm < 1 Or prevAdm > MAX_ADMIN_LEVEL Then _
        Err.Raise 5, "GeoModule", _
                  "The previous admin level of " & tabId & " reads " & prevAdm

    'The caller fires on every Validate with no idea whether the level
    'changed.
    If prevAdm = actAdm Then
        LinelistEventsManager.LLExitBusyState
        Exit Sub
    End If

    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))

    pass.UnProtect "_active"
    unprotected = True

    'One read of the whole header row; the match is a string test in memory.
    headerTexts = HeaderFormulaTexts(headerRng)

    For counter = 1 To headerRng.Columns.Count
        If InStr(1, CStr(headerTexts(1, counter)), rngName) > 0 Then
            Set valuesRng = Nothing
            'headerCellName is a String, so without the reset it keeps the
            'previous column's name when the read below raises -- a header
            'cell has a name only when something named it -- and the Replace
            'would then resolve the previous column's VALUES block.
            headerCellName = vbNullString

            On Error Resume Next
            headerCellName = headerRng.Cells(1, counter).Name.Name
            If LenB(headerCellName) > 0 Then
                Set valuesRng = sh.Range(Replace(headerCellName, "LABEL", "VALUES"))
            End If
            On Error GoTo ErrSPT

            If Not valuesRng Is Nothing Then
                'The named block ends two rows above the Total and Missing
                'rows of its own table -- RowsCategoriesRange trims the two
                'footer rows by count -- and their formulas move levels with
                'the block, so the range is grown to reach them.
                Set valuesRng = sh.Range(valuesRng.Cells(1, 1), _
                                         valuesRng.Cells(valuesRng.Rows.Count + 2, 1))

                sp.RewriteFormulas valuesRng, "concat_adm" & prevAdm, _
                                   "concat_adm" & actAdm
            End If
        End If
    Next

    sh.Range(rngName).Offset(, 1).Value = actAdm
    sh.UsedRange.Calculate

    pass.Protect sh, allowShapes:=True
    LinelistEventsManager.LLExitBusyState
    Exit Sub

ErrSPT:
    'The protection is put back only when this run took it off, so a raise
    'above the UnProtect leaves a deliberately open sheet open.
    If unprotected Then pass.Protect sh, allowShapes:=True
    LinelistEventsManager.LLExitBusyState
    ReportGeoError Err.Description
End Sub

' @description Read the header row of formulas in one crossing. A one-cell
'              range answers a scalar, so it is wrapped to give the caller one
'              shape.
' @param headerRng The SPT_FORMULA_COLUMN_ row of one section
Private Function HeaderFormulaTexts(ByVal headerRng As Range) As Variant
    Dim oneCell(1 To 1, 1 To 1) As Variant

    If headerRng.Cells.Count = 1 Then
        oneCell(1, 1) = headerRng.Formula
        HeaderFormulaTexts = oneCell
    Else
        HeaderFormulaTexts = headerRng.Formula
    End If
End Function
