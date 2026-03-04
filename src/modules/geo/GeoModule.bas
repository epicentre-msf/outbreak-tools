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
Private Const LLSHEET As String = "LinelistTranslation"
Private Const SPATIALSHEET As String = "spatial_tables__"
Private Const PASSSHEET As String = "__pass"

'@section Module-Level State
'===============================================================================

Private historicGeoTable As BetterArray
Private historicHFTable As BetterArray
Private concatenateGeoTable As BetterArray
Private concatenateHFTable As BetterArray
Private geo As ILLGeo
Private drop As IDropdownLists
Private tradmess As ITranslationObject
Private lltrads As ILLTranslation
Private pass As IPasswords

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

    Set lltrads = LLTranslation.Create(wb.Worksheets(LLSHEET))
    Set tradmess = lltrads.TransObject()
End Sub

' @description Initialize passwords and translations for spatial analysis events.
Private Sub InitializeSpatialElements()
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set pass = Passwords.Create(wb.Worksheets(PASSSHEET))
End Sub

'@section Application State
'===============================================================================

' @description Suspend heavy Excel UI features for performance.
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
    Application.cursor = cursor
End Sub

' @description Restore Excel UI to normal state.
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
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
    BusyApp

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

    NotBusyApp

    F_Geo.TXT_Msg.Value = vbNullString
    F_Geo.Show
    Exit Sub

ErrLoadGeo:
    MsgBox tradmess.TranslatedValue("MSG_ErrGeo"), _
           vbOKOnly + vbCritical, _
           tradmess.TranslatedValue("MSG_Error")
    NotBusyApp
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
                          Optional ByVal scope As Byte = 0)

    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm2.Clear
        F_Geo.LST_Adm3.Clear
        F_Geo.LST_Adm4.Clear
    Else
        F_Geo.LST_AdmF2.Clear
        F_Geo.LST_AdmF3.Clear
        F_Geo.LST_AdmF4.Clear
    End If

    Dim adminTable As BetterArray
    Application.cursor = xlNorthwestArrow

    Set adminTable = geo.GeoLevel(LevelAdmin2, scope, selectedAdmin1)
    F_Geo.TXT_Msg.Value = selectedAdmin1

    If adminTable.Length = 0 Then Exit Sub

    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm2.List = adminTable.Items
    Else
        F_Geo.LST_AdmF2.List = adminTable.Items
    End If

    Application.cursor = xlDefault
End Sub

' @description Show admin3 list filtered by selected admin1 and admin2.
'@EntryPoint
Public Sub ShowAdmin3List(ByVal selectedAdmin2 As String, _
                          Optional ByVal scope As Byte = 0, _
                          Optional ByVal separator As String = " | ")

    Dim selectedAdmin1 As String
    Dim concatenateAdmins As String
    Dim adminTable As BetterArray
    Dim adminNames As BetterArray

    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm3.Clear
        F_Geo.LST_Adm4.Clear
        selectedAdmin1 = F_Geo.LST_Adm1.Value
        concatenateAdmins = selectedAdmin1 & separator & selectedAdmin2
    Else
        F_Geo.LST_AdmF3.Clear
        F_Geo.LST_AdmF4.Clear
        selectedAdmin1 = F_Geo.LST_AdmF1.Value
        concatenateAdmins = selectedAdmin2 & separator & selectedAdmin1
    End If

    Set adminNames = New BetterArray
    adminNames.LowerBound = 1
    Application.cursor = xlNorthwestArrow

    adminNames.Push selectedAdmin1, selectedAdmin2
    Set adminTable = geo.GeoLevel(LevelAdmin3, scope, adminNames)
    F_Geo.TXT_Msg.Value = concatenateAdmins

    If adminTable.Length = 0 Then Exit Sub

    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm3.List = adminTable.Items
    Else
        F_Geo.LST_AdmF3.List = adminTable.Items
    End If

    Application.cursor = xlDefault
End Sub

' @description Show admin4 list filtered by selected admin1, admin2, and admin3.
'@EntryPoint
Public Sub ShowAdmin4List(ByVal selectedAdmin3 As String, _
                          Optional ByVal scope As Byte = 0, _
                          Optional ByVal separator As String = " | ")

    Dim adminTable As BetterArray
    Dim adminNames As BetterArray
    Dim selectedAdmin1 As String
    Dim selectedAdmin2 As String
    Dim concatenateAdmins As String

    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm4.Clear
        selectedAdmin1 = F_Geo.LST_Adm1.Value
        selectedAdmin2 = F_Geo.LST_Adm2.Value
        concatenateAdmins = selectedAdmin1 & separator & _
                            selectedAdmin2 & separator & _
                            selectedAdmin3
    Else
        F_Geo.LST_AdmF4.Clear
        selectedAdmin1 = F_Geo.LST_AdmF1.Value
        selectedAdmin2 = F_Geo.LST_AdmF2.Value
        concatenateAdmins = selectedAdmin3 & separator & _
                            selectedAdmin2 & separator & _
                            selectedAdmin1
    End If

    Set adminNames = New BetterArray
    adminNames.LowerBound = 1
    adminNames.Push selectedAdmin1, selectedAdmin2, selectedAdmin3

    Application.cursor = xlNorthwestArrow

    Set adminTable = geo.GeoLevel(LevelAdmin4, GeoScopeAdmin, adminNames)
    F_Geo.TXT_Msg.Value = concatenateAdmins

    If adminTable.Length = 0 Then Exit Sub

    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm4.List = adminTable.Items
    Else
        F_Geo.LST_AdmF4.List = adminTable.Items
    End If

    Application.cursor = xlDefault
End Sub

'@section Spatial Table Updates
'===============================================================================

' @description Update all spatial tables from HList filtered data.
'@EntryPoint
Public Sub UpdateSpTables()
    Dim sp As ILLSpatial
    Set sp = LLSpatial.Create(ThisWorkbook.Worksheets(SPATIALSHEET))

    UpdateFilterTables calculate:=False

    BusyApp
    sp.Update
    NotBusyApp
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

    BusyApp cursor:=xlNorthwestArrow
    InitializeSpatialElements

    On Error GoTo ErrSPT

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

ErrSPT:
    pass.Protect sh, True
    NotBusyApp
End Sub
