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

    On Error GoTo ErrManage

    Dim tradssh As Worksheet
    Dim geosh As Worksheet
    Dim passsh As Worksheet
    Dim geo As ILLGeo

    BusyApp

    With ThisWorkbook
        Set tradssh = .Worksheets("LinelistTranslation")
        Set geosh = .Worksheets("Geo")
        Set passsh = .Worksheets("__pass")
    End With

    'Language of forms in the dictionary changes
    If Not (Intersect(Target, Me.Range("RNG_LLForm")) Is Nothing) Then

        'Language of LinelistForms
        tradssh.Range("RNG_LLLanguage").Value = Target.Value
        tradssh.calculate

        'Language Code in the Geo Sheet
        Set geo = LLGeo.Create(geosh)
        geo.Translate rawNames:=True
        geosh.Range("RNG_GeoLangCode").Value = tradssh.Range("RNG_LLLanguageCode").Value
        geosh.calculate

    'password changes
    ElseIf Not (Intersect(Target, Me.Range("RNG_LLPassword")) Is Nothing) Then

        passsh.Range("RNG_DebuggingPassword").Value = Target.Value

    'Language of the setup changes (langage of elements in  the linelist)
    ElseIf Not (Intersect(Target, Me.Range("RNG_LangSetup")) Is Nothing) Then
        tradssh.Range("RNG_DictionaryLanguage").Value = Target.Value
        geosh.Range("RNG_MetaLang").Value = Target.Value
    End If

ErrManage:
    NotBusyApp
End Sub
