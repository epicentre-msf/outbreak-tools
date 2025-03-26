Attribute VB_Name = "EventsMainSheet"
Option Explicit

'speed app
Private Sub BusyApp()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)


    Dim tradssh As Worksheet
    Dim geosh As Worksheet
    Dim passsh As Worksheet
    Dim designsh As Worksheet
    Dim geo As ILLGeo
    Dim rngName As String

    BusyApp

    With ThisWorkbook
        Set tradssh = .Worksheets("LinelistTranslation")
        Set geosh = .Worksheets("Geo")
        Set passsh = .Worksheets("__pass")
        Set designsh = .Worksheets("LinelistStyle")
    End With

    On Error Resume Next
        rngName = Target.Name.Name
    On Error GoTo ErrManage

    Select Case rngName
    'Language of forms in the dictionary changes
    Case "RNG_LLForm"

        'Language of LinelistForms
        tradssh.Range("RNG_LLLanguage").Value = Target.Value
        tradssh.calculate

        'Language Code in the Geo Sheet
        Set geo = LLGeo.Create(geosh)
        geo.Translate rawNames:=True
        geosh.Range("RNG_GeoLangCode").Value = tradssh.Range("RNG_LLLanguageCode").Value
        geosh.calculate

    'password changes
    Case "RNG_LLPassword"

        passsh.Range("RNG_DebuggingPassword").Value = Target.Value

    'Language of the setup changes (langage of elements in  the linelist)
    Case "RNG_LangSetup"
        tradssh.Range("RNG_DictionaryLanguage").Value = Target.Value
        geosh.Range("RNG_MetaLang").Value = Target.Value
    
    'Design change
    Case "RNG_DesignLL"
        designsh.Range("DESIGNTYPE").Value = Target.Value        
    End Select

ErrManage:
    NotBusyApp
End Sub
