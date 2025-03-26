Attribute VB_Name = "LinelistGeo"
'@IgnoreModule ImplicitActiveSheetReference

Option Explicit
Option Base 1
Option Private Module

Private Const GEOSHEET As String = "Geo"
Private Const DROPDOWNSHEET As String = "dropdown_lists__"
Private Const TRADSHEET As String = "Translations"
Private Const LLSHEET As String = "LinelistTranslation"

Private historicGeoTable As BetterArray                    'Historic of Geo
Private historicHFTable As BetterArray                     'Historic of health facility
Private concatenateGeoTable As BetterArray
Private concatenateHFTable As BetterArray
Private geo As ILLGeo
Private drop As IDropdownLists
Private tradmess As ITranslation
Private lltrads As ILLTranslations

'Initialize some elements
Private Sub InitializeElements()
    Dim wb As Workbook

    Set historicGeoTable = New BetterArray
    Set historicHFTable = New BetterArray

    Set concatenateGeoTable = New BetterArray
    Set concatenateHFTable = New BetterArray

    Set wb = ThisWorkbook
    Set geo = LLGeo.Create(wb.Worksheets(GEOSHEET))
    Set drop = DropdownLists.Create(wb.Worksheets(DROPDOWNSHEET))

     Set lltrads = LLTranslations.Create( _
                                        wb.Worksheets(LLSHEET), _
                                        wb.Worksheets(TRADSHEET) _
                                        )
    Set tradmess = lltrads.TransObject()
End Sub

'Speed up before a work
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableAnimations = False
End Sub

'Return previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableAnimations = True
End Sub


'Load the form for geoHelper
'Facility: hfOrGeo = 1, Geo: hfOrGeo = 0
Public Sub LoadGeo(ByVal hfOrGeo As Byte)

    Dim transValue As BetterArray

    On Error GoTo ErrLoadGeo

    InitializeElements

    Set transValue = New BetterArray
    BusyApp
    
    Select Case hfOrGeo
        
    Case GeoScopeAdmin 'Geo
        
        'Add Caption for  each adminstrative leveles in the form
        F_Geo.LBL_Adm1.Caption = geo.GeoNames("adm1_name")
        F_Geo.LBL_Adm2.Caption = geo.GeoNames("adm2_name")
        F_Geo.LBL_Adm3.Caption = geo.GeoNames("adm3_name")
        F_Geo.LBL_Adm4.Caption = geo.GeoNames("adm4_name")

        drop.ClearList "admin2"
        drop.ClearList "admin3"
        drop.ClearList "admin4"

        'Before doing the whole all thing, we need to test if the T_Adm data is empty or not
        If (Not geo.HasNoData()) Then
            
            'Update admin 1
            Set transValue = geo.GeoLevel(LevelAdmin1, GeoScopeAdmin)
            
            'Clear selected values on Geo
            ClearLists
            F_Geo.LST_Adm1.List = transValue.Items
            
            'There is a range named adm4_concat in the workbook
            concatenateGeoTable.FromExcelRange Range("adm4_concat")
            F_Geo.LST_ListeAgre.List = concatenateGeoTable.Items

        End If

        'Historic for geographic data and facility data
        historicGeoTable.FromExcelRange Range("histo_geo")
        F_Geo.LST_Histo.List = historicGeoTable.Items

        F_Geo.FRM_Facility.Visible = False
        F_Geo.FRM_Geo.Visible = True
        F_Geo.LBL_Fac1.Visible = False
        F_Geo.LBL_Geo1.Visible = True

    Case GeoScopeHF 'HF

        'Adding caption for each admnistrative levels
        F_Geo.LBL_Adm4F.Caption = geo.GeoNames("hf_name")
        F_Geo.LBL_Adm3F.Caption = geo.GeoNames("adm3_name")
        F_Geo.LBL_Adm2F.Caption = geo.GeoNames("adm2_name")
        F_Geo.LBL_Adm1F.Caption = geo.GeoNames("adm1_name")

        If (Not geo.HasNoData()) Then
            Set transValue = geo.GeoLevel(LevelAdmin1, GeoScopeHF)
            
            'Clear selected values on health facility
            ClearLists
            F_Geo.LST_AdmF1.List = transValue.Items
            'There is a range named hf_concat in the workbook
            concatenateHFTable.FromExcelRange Range("hf_concat")
            F_Geo.LST_ListeAgreF.List = concatenateHFTable.Items
        End If

        'Historic HF
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

Private Sub ClearLists()
    Dim counter As Integer

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


'Show the list of admin 2 for geo frame, given a selected admin 1
'@EntryPoint
Public Sub ShowAdmin2List(ByVal selectedAdmin1 As String, _
                          Optional ByVal scope As Byte = 0)
    
    If scope = GeoScopeAdmin Then
        'Clear Geo
        F_Geo.LST_Adm2.Clear
        F_Geo.LST_Adm3.Clear
        F_Geo.LST_Adm4.Clear
    Else
        'Clear Facilities
        F_Geo.LST_AdmF2.Clear
        F_Geo.LST_AdmF3.Clear
        F_Geo.LST_AdmF4.Clear
    End If

    Dim adminTable As BetterArray
    Set adminTable = New BetterArray
    Application.cursor = xlNorthwestArrow

    'Search if the value exists in the 2 dimensional table T_Adm1 previously initialized
    Set adminTable = geo.GeoLevel(LevelAdmin2, scope, selectedAdmin1)
    F_Geo.TXT_Msg.Value = selectedAdmin1

    'update if only next level is available
    If adminTable.Length = 0 Then Exit Sub
    
    If scope = GeoScopeAdmin Then
        F_Geo.LST_Adm2.List = adminTable.Items
    Else
        F_Geo.LST_AdmF2.List = adminTable.Items
    End If

    Application.cursor = xlDefault
End Sub

'Show the third list for the geobase given selected admin1 and admin2
'@EntryPoint
Public Sub ShowAdmin3List(ByVal selectedAdmin2 As String, _
                          Optional ByVal scope As Byte = 0, _
                          Optional ByVal separator As String = " | ")

    Dim selectedAdmin1 As String
    Dim concatenateAdmins As String
    Dim adminTable As BetterArray 'The table or corresponding admin 3
    Dim adminNames As BetterArray 'The table of selected admin

    'Clear remaining admin levels
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
    'It is important to have 1 as lowerbound for geoLevel function
    adminNames.LowerBound = 1

    Application.cursor = xlNorthwestArrow

    'add admin 1 and admin 2 for filtering
    adminNames.Push selectedAdmin1, selectedAdmin2

    'Just filter and show
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

'Show the fourth list for the geobase given selected admin1, admin2, and admin3
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

