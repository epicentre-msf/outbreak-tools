Attribute VB_Name = "GeoModule"
Attribute VB_Description = "Combined geo and spatial analysis logic for the linelist"

'@Folder("Geo")
'@ModuleDescription("Combined geo and spatial analysis logic for the linelist")
'@IgnoreModule UnrecognizedAnnotation, ImplicitActiveSheetReference, UseMeaningfulName, HungarianNotation

Option Explicit
Option Base 1
Option Private Module

'@section Constants
'===============================================================================

Private Const GEOSHEET As String = "Geo"
Private Const DROPDOWNSHEET As String = "dropdown_lists__"
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const PASSSHEET As String = "__pass"

'@section Module-Level State
'===============================================================================

Private historicGeoTable As BetterArray
Private historicHFTable As BetterArray
Private concatenateGeoTable As BetterArray
Private concatenateHFTable As BetterArray
Private geo As LLGeo
Private drop As DropdownLists
Private pass As Passwords

'@section Initialization
'===============================================================================

' @description Initialize geo elements: LLGeo, dropdowns, and translations.
Private Sub InitializeGeoElements()
    Dim wb As Workbook

    Set historicGeoTable = New BetterArray
    Set historicHFTable = New BetterArray
    Set concatenateGeoTable = New BetterArray
    Set concatenateHFTable = New BetterArray

    Set wb = ThisWorkbook
    Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))
    Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))
End Sub

' @description Initialize passwords and translations for spatial analysis events.
Private Sub InitializeSpatialElements()
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEET))
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
    Dim transValue As BetterArray

    On Error GoTo ErrLoadGeo

    InitializeGeoElements

    Set transValue = New BetterArray
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow

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
            Set transValue = geo.GeoLevel(LevelAdmin1, GeoScopeAdmin)
            ClearLists
            F_Geo.LST_Adm1.List = transValue.Items
            concatenateGeoTable.FromExcelRange Range("adm4_concat")
            F_Geo.LST_ListeAgre.List = concatenateGeoTable.Items
        End If

        historicGeoTable.FromExcelRange Range("histo_geo")
        F_Geo.LST_Histo.List = historicGeoTable.Items

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
            Set transValue = geo.GeoLevel(LevelAdmin1, GeoScopeHF)
            ClearLists
            F_Geo.LST_AdmF1.List = transValue.Items
            concatenateHFTable.FromExcelRange Range("hf_concat")
            F_Geo.LST_ListeAgreF.List = concatenateHFTable.Items
        End If

        historicHFTable.FromExcelRange Range("histo_hf")
        F_Geo.LST_HistoF.List = historicHFTable.Items
        F_Geo.FRM_Facility.Visible = True
        F_Geo.FRM_Geo.Visible = False
        F_Geo.LBL_Fac1.Visible = True
        F_Geo.LBL_Geo1.Visible = False

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

' @description Clear all list controls in the F_Geo form.
Private Sub ClearLists()
    Dim counter As Long

    With F_Geo
        .LST_AdmF1.Value = ""
        .LST_Adm1.Value = ""
        .LST_ListeAgreF.Value = ""
        .LST_ListeAgre.Value = ""
        .LST_Histo.Value = ""
        .LST_HistoF.Value = ""
        For counter = 2 To 4
            .Controls("LST_Adm" & counter).Clear
            .Controls("LST_AdmF" & counter).Clear
        Next
    End With
End Sub

'@section Admin Cascade — ShowAdmin*List
'===============================================================================

' @description Show admin2 list filtered by selected admin1.
'@EntryPoint
Public Sub ShowAdmin2List(ByVal selectedAdmin1 As String, _
                          Optional ByVal scope As Byte = GeoScopeAdmin)
    Dim adminTable As BetterArray

    On Error GoTo ErrShowAdmin2
    Application.Cursor = xlNorthwestArrow

    With F_Geo
        If scope = GeoScopeAdmin Then
            .LST_Adm2.Clear
            .LST_Adm3.Clear
            .LST_Adm4.Clear
        Else
            .LST_AdmF2.Clear
            .LST_AdmF3.Clear
            .LST_AdmF4.Clear
        End If

        Set adminTable = geo.GeoLevel(LevelAdmin2, scope, selectedAdmin1)
        .TXT_Msg.Value = selectedAdmin1

        If adminTable.Length > 0 Then
            If scope = GeoScopeAdmin Then
                .LST_Adm2.List = adminTable.Items
            Else
                .LST_AdmF2.List = adminTable.Items
            End If
        End If
    End With

    Application.Cursor = xlDefault
    Exit Sub

ErrShowAdmin2:
    Application.Cursor = xlDefault
    ReportGeoError Err.Description
End Sub

' @description Show admin3 list filtered by selected admin1 and admin2.
'@EntryPoint
Public Sub ShowAdmin3List(ByVal selectedAdmin2 As String, _
                          Optional ByVal scope As Byte = GeoScopeAdmin, _
                          Optional ByVal separator As String = " | ")

    Dim selectedAdmin1 As String
    Dim concatenateAdmins As String
    Dim adminTable As BetterArray
    Dim adminNames As BetterArray

    On Error GoTo ErrShowAdmin3
    Application.Cursor = xlNorthwestArrow

    With F_Geo
        If scope = GeoScopeAdmin Then
            .LST_Adm3.Clear
            .LST_Adm4.Clear
            selectedAdmin1 = .LST_Adm1.Value
            concatenateAdmins = selectedAdmin1 & separator & selectedAdmin2
        Else
            .LST_AdmF3.Clear
            .LST_AdmF4.Clear
            selectedAdmin1 = .LST_AdmF1.Value
            concatenateAdmins = selectedAdmin2 & separator & selectedAdmin1
        End If

        Set adminNames = New BetterArray
        adminNames.LowerBound = 1
        adminNames.Push selectedAdmin1, selectedAdmin2

        Set adminTable = geo.GeoLevel(LevelAdmin3, scope, adminNames)
        .TXT_Msg.Value = concatenateAdmins

        If adminTable.Length > 0 Then
            If scope = GeoScopeAdmin Then
                .LST_Adm3.List = adminTable.Items
            Else
                .LST_AdmF3.List = adminTable.Items
            End If
        End If
    End With

    Application.Cursor = xlDefault
    Exit Sub

ErrShowAdmin3:
    Application.Cursor = xlDefault
    ReportGeoError Err.Description
End Sub

' @description Show admin4 list filtered by selected admin1, admin2, and admin3.
'@EntryPoint
Public Sub ShowAdmin4List(ByVal selectedAdmin3 As String, _
                          Optional ByVal scope As Byte = GeoScopeAdmin, _
                          Optional ByVal separator As String = " | ")

    Dim adminTable As BetterArray
    Dim adminNames As BetterArray
    Dim selectedAdmin1 As String
    Dim selectedAdmin2 As String
    Dim concatenateAdmins As String

    On Error GoTo ErrShowAdmin4
    Application.Cursor = xlNorthwestArrow

    With F_Geo
        If scope = GeoScopeAdmin Then
            .LST_Adm4.Clear
            selectedAdmin1 = .LST_Adm1.Value
            selectedAdmin2 = .LST_Adm2.Value
            concatenateAdmins = selectedAdmin1 & separator & _
                                selectedAdmin2 & separator & _
                                selectedAdmin3
        Else
            .LST_AdmF4.Clear
            selectedAdmin1 = .LST_AdmF1.Value
            selectedAdmin2 = .LST_AdmF2.Value
            concatenateAdmins = selectedAdmin3 & separator & _
                                selectedAdmin2 & separator & _
                                selectedAdmin1
        End If

        Set adminNames = New BetterArray
        adminNames.LowerBound = 1
        adminNames.Push selectedAdmin1, selectedAdmin2, selectedAdmin3

        Set adminTable = geo.GeoLevel(LevelAdmin4, scope, adminNames)
        .TXT_Msg.Value = concatenateAdmins

        If adminTable.Length > 0 Then
            If scope = GeoScopeAdmin Then
                .LST_Adm4.List = adminTable.Items
            Else
                .LST_AdmF4.List = adminTable.Items
            End If
        End If
    End With

    Application.Cursor = xlDefault
    Exit Sub

ErrShowAdmin4:
    Application.Cursor = xlDefault
    ReportGeoError Err.Description
End Sub

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
' @param rngName Named range of the admin level selector
' @param actAdm New admin level (number of admin levels selected)
'@EntryPoint
Public Sub UpdateSpatioTemporalFormulas(ByVal rngName As String, _
                                        ByVal actAdm As Long)
    Dim tabId As String
    Dim prevAdm As Long
    Dim sh As Worksheet
    Dim counter As Long
    Dim headerRng As Range
    Dim cellRng As Range
    Dim valuesRng As Range
    Dim headerFormula As String
    Dim valuesFormula As String
    Dim headerCellName As String
    Dim hasFormula As Boolean

    'The handler is armed first, so a raise in the busy-state entry or in
    'InitializeSpatialElements reaches ErrSPT and restores the application.
    On Error GoTo ErrSPT
    LinelistEventsManager.LLEnterBusyState busyCursor:=xlNorthwestArrow
    InitializeSpatialElements

    Set sh = ActiveSheet
    tabId = "SPT_" & Split(rngName, "_")(3)
    Set headerRng = sh.Range("SPT_FORMULA_COLUMN_" & tabId)
    prevAdm = sh.Range(rngName).Offset(, 1).Value

    pass.UnProtect "_active"

    For counter = 1 To headerRng.Columns.Count
        headerFormula = Replace(headerRng.Cells(1, counter).Formula, "=", vbNullString)
        headerFormula = Application.WorksheetFunction.Trim(headerFormula)

        If InStr(1, headerFormula, rngName) > 0 Then
            Set valuesRng = Nothing

            On Error Resume Next
            headerCellName = headerRng.Cells(1, counter).Name.Name
            Set valuesRng = sh.Range(Replace(headerCellName, "LABEL", "VALUES"))
            On Error GoTo ErrSPT

            If Not valuesRng Is Nothing Then
                Set valuesRng = sh.Range(valuesRng.Cells(1, 1), _
                                         valuesRng.Cells(valuesRng.Rows.Count + 2, 1))

                For Each cellRng In valuesRng
                    hasFormula = False
                    valuesFormula = cellRng.FormulaArray

                    If valuesFormula = vbNullString Then
                        valuesFormula = cellRng.Formula
                        hasFormula = True
                    End If

                    If InStr(1, valuesFormula, "concat_adm" & prevAdm) > 0 Then
                        valuesFormula = Replace(valuesFormula, _
                                                "concat_adm" & prevAdm, _
                                                "concat_adm" & actAdm)

                        If hasFormula Then
                            cellRng.Formula = valuesFormula
                        Else
                            cellRng.FormulaArray = valuesFormula
                        End If
                    End If
                Next
            End If
        End If
    Next

    sh.Range(rngName).Offset(, 1).Value = actAdm
    sh.UsedRange.Calculate

    pass.Protect sh, True
    LinelistEventsManager.LLExitBusyState
    Exit Sub

ErrSPT:
    'sh is Nothing when the raise came before the sheet was read, and nothing
    'was unprotected on that path.
    If Not sh Is Nothing Then pass.Protect sh, True
    LinelistEventsManager.LLExitBusyState
    ReportGeoError Err.Description
End Sub
