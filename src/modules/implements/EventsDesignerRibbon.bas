Attribute VB_Name = "EventsDesignerRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the designer"
Option Explicit
Option Private Module

'@Folder("Designer Events")
'@ModuleDescription("Events associated to the Ribbon Menu in the designer")

'Designer Translation sheet name
Private Const DESIGNERTRADSHEET As String = "DesignerTranslation"
'Linelist translation sheet name
Private Const LINELISTTRADSHEET As String = "LinelistTranslation"
'Designer main sheet name
Private Const DESIGNERMAINSHEET As String = "Main"
'All the ribbon object Ribbon
Private ribbonUI As IRibbonUI

'speed up process
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

'@Description("Callback when the button loaded")
'@EntryPoint
Public Sub ribbonLoaded(ByRef ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub

'@Description("Triggers event to update all the labels by relaunching all the callbacks")
Private Sub UpdateLabels
    ribbonUI.Invalidate
End Sub

'@Description("Callback for getLabel (Depending on the language)")
'@EntryPoint
Public Sub LangLabel(control As IRibbonControl, ByRef returnedVal)
    Attribute LangLabel.VB_Description = "Callback for getLabel (Depending on the language)"

    Dim desTrads As IDesTranslation
    Dim codeId As String
    Dim tradsh As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets("DesignerTranslation")
    Set desTrads = DesTranslation.Create(tradsh)
    codeId = control.Id

    returnedVal = desTrads.TranslationMsg(codeId)
End Sub

'@Description("Callback for btnDelGeo onAction: Delete the geobase")
'@EntryPoint
Public Sub clickDelGeo(control As IRibbonControl)
Attribute clickDelGeo.VB_Description = "Callback for btnDelGeo onAction: Delete the geobase"
    Dim geosh As Worksheet
    Dim geo As ILLGeo
    Dim wb As Workbook

    On Error GoTo ErrGeo
    BusyApp

    Set wb = ThisWorkbook
    Set geosh = wb.Worksheets("Geo")
    Set geo = LLGeo.Create(geosh)

    'Clear the geobase data
    geo.Clear

ErrGeo:
    NotBusyApp
End Sub

'@Description("Callback for btnClear onAction": Clear the entries)
'@EntryPoint
Public Sub clickClearEnt(control As IRibbonControl)

    Dim wb As Workbook
    Dim mainsh As Worksheet
    Dim mainobj As IMain

    BusyApp

    On Error GoTo ErrEnt

    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set mainobj = Main.Create(mainsh)
    mainobj.ClearInputRanges clearValues := True

ErrEnt:
    NotBusyApp
End Sub

'@Description("Callback for btnTransAdd onAction: Import Linelist translations")
'@EntryPoint
Public Sub clickImpTrans(control As IRibbonControl)
Attribute clickImpTrans.VB_Description = "Callback for btnTransAdd onAction: Import Linelist translations"


    Dim io As IOSFiles
    Dim wb As Workbook 'Actual workbook
    Dim impsh As Worksheet 'Imported worksheet
    Dim impwb As Workbook 'Imported workbook
    Dim actsh As Worksheet 'Actual worksheet
    Dim actLo As ListObject 'Actual ListObject
    Dim impLo As ListObject 'Imported ListObject
    Dim actcsTab As ICustomTable 'Actual custom table
    Dim impcsTab As ICustomTable 'Imported custom table
    Dim loListName As BetterArray 'List of listObjects to import
    Dim tradsSheetsList As BetterArray 'Listof sheets to import
    Dim counter As Long
    Dim sheetName As String

    Set wb = ThisWorkbook

    'Import the translations for
    Set io = OSFiles.Create()
    Set loListName = New BetterArray
    Set tradsSheetsList = New BetterArray

    io.LoadFile "*.xlsb"
    If io.HasValidFile() Then
        BusyApp

        tradsSheetsList.Push LINELISTTRADSHEET, DESIGNERTRADSHEET
        loListName.Push "T_TradLLShapes", "T_TradLLMsg", "T_TradLLForms", "T_TradLLRibbon", _
                        "T_tradMsg", "T_tradRange", "T_tradShape"
        Set impwb = Workbooks.Open(io.File())

        For counter = tradsSheetsList.LowerBound To tradsSheetsList.UpperBound
            sheetName = tradsSheetsList.Item(counter)
            Set actsh = wb.Worksheets(sheetName)
            On Error GoTo ExitTrads
            Set impsh = impwb.Worksheets(sheetName)
            For Each actLo In actsh.ListObjects
                If loListName.includes(actLo.Name) Then
                    Set actcsTab = CustomTable.Create(actLo)
                    Set impLo = impsh.ListObjects(actLo.Name)
                    Set impcsTab = CustomTable.Create(impLo)
                    actcsTab.Import impcsTab
                End If
            Next
            actsh.Calculate
        Next
        On Error GoTo 0
    End If
ExitTrads:
    On Error Resume Next
    impwb.Close saveChanges:=False
    NotBusyApp
    MsgBox "Done!"
    On Error GoTo 0
End Sub

'@Description("Callback for langDrop onAction: Change the language of the designer")
'@EntryPoint
Public Sub clickLangChange(control As IRibbonControl, langId As String, Index As Integer)
    Attribute clickLangChange.VB_Description = "Callback for langDrop onAction: Change the language of the designer"

    'Language code in the designer worksheet
    Const RNGLANGCODE As String = "RNG_MainLangCode"

    'langId is the language code
    Dim tradsh As Worksheet
    Dim desTrads As IDesTranslation
    Dim mainsh As Worksheet
    Dim wb As Workbook

    BusyApp

    On Error GoTo ExitLang

    Set wb = ThisWorkbook
    Set mainsh  = wb.Worksheets("Main")
    Set tradsh = wb.Worksheets("DesignerTranslation")
    Set desTrads = DesTranslation.Create(tradsh)

    tradsh.Range(RNGLANGCODE).Value = langId
    tradsh.Calculate
    desTrads.TranslateDesigner mainsh

    'Update all the labels on the ribbon
    UpdateLabels

ExitLang:
    NotBusyApp
End Sub

'@Description("Callback for btnOpen onAction: Open another linelist file")
'@EntryPoint
Public Sub clickOpen(control As IRibbonControl)
Attribute clickOpen.VB_Description = "Callback for btnOpen onAction: Open another linelist file"

    Dim io As IOSFiles
    Dim trads As IDesTranslation
    Set io = OSFiles.Create()

    io.LoadFile "*.xlsb"                         '
    If Not io.HasValidFile Then Exit Sub

    On Error GoTo ErrorManage
    NotBusyApp
    Application.Workbooks.Open FileName:=io.File(), ReadOnly:=False
    Exit Sub

ErrorManage:
    On Error Resume Next
    Set trads = DesTranslation.Create(ThisWorkbook.Worksheets(DESIGNERTRADSHEET))
    MsgBox trads.TranslationMsg("MSG_TitlePassWord"), vbCritical, _
    trads.TranslationMsg("MSG_PassWord")
    On Error GoTo 0
End Sub
