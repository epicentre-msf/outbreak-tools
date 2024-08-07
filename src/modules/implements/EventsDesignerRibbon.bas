Attribute VB_Name = "EventsDesignerRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the designer"
Option Explicit
Option Private Module
'@Folder("Designer Events")
'@IgnoreModule ParameterNotUsed
'@ModuleDescription("Events associated to the Ribbon Menu in the designer")

'Designer Translation sheet name
Private Const DESIGNERTRADSHEET As String = "DesignerTranslation"
'Linelist translation sheet name
Private Const LINELISTTRADSHEET As String = "LinelistTranslation"
'Designer main sheet name
Private Const DESIGNERMAINSHEET As String = "Main"
Private Const PASSWORDSHEET As String = "__pass"
'Linelist Style worksheet
Private Const FORMATSHEET As String = "LinelistStyle"
'All the ribbon object Ribbon
Private ribbonUI As IRibbonUI

'speed up process
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlNorthwestArrow)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.cursor = cursor
    Application.Calculation = xlCalculationManual
End Sub

Private Sub NotBusyApp()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub

'@Description("Callback when the button loaded")
'@EntryPoint
Public Sub ribbonLoaded(ByRef ribbon As IRibbonUI)
    Set ribbonUI = ribbon
End Sub

'Triggers event to update all the labels by relaunching all the callbacks
Private Sub UpdateLabels()
    ribbonUI.Invalidate
End Sub

'@Description("Callback for getLabel (Depending on the language)")
'@EntryPoint
Public Sub LangLabel(control As IRibbonControl, ByRef returnedVal)

    Dim desTrads As IDesTranslation
    Dim codeId As String
    Dim tradsh As Worksheet
    Dim wb As Workbook

    Set wb = ThisWorkbook
    Set tradsh = wb.Worksheets("DesignerTranslation")
    Set desTrads = DesTranslation.Create(tradsh)
    codeId = control.ID

    returnedVal = desTrads.TranslationMsg(codeId)
End Sub

'@Description("Callback for btnDelGeo onAction: Delete the geobase")
'@EntryPoint
Public Sub clickDelGeo(control As IRibbonControl)
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
    mainobj.ClearInputRanges clearValues:=True

ErrEnt:
    NotBusyApp
End Sub

'@Description("Callback for btnTransAdd onAction: Import Linelist translations")
'@EntryPoint
Public Sub clickImpTrans(control As IRibbonControl)


    Dim io As IOSFiles
    Dim wb As Workbook 'Actual workbook
    Dim impsh As Worksheet 'Imported worksheet
    Dim impwb As Workbook 'Imported workbook
    Dim actsh As Worksheet 'Actual worksheet
    Dim actLo As listObject 'Actual ListObject
    Dim impLo As listObject 'Imported ListObject
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

    io.LoadFile "*.xlsx"
    If io.HasValidFile() Then
        BusyApp

        tradsSheetsList.Push LINELISTTRADSHEET, DESIGNERTRADSHEET
        loListName.Push "t_tradllshapes", "t_tradllmsg", "t_tradllforms", "t_tradllribbon", _
                        "t_tradmsg", "t_tradrange", "t_tradshape"
        Set impwb = Workbooks.Open(io.File())

        For counter = tradsSheetsList.LowerBound To tradsSheetsList.UpperBound
            sheetName = tradsSheetsList.Item(counter)
            Set actsh = wb.Worksheets(sheetName)
            On Error GoTo ExitTrads
            Set impsh = impwb.Worksheets(sheetName)
            For Each actLo In actsh.ListObjects
                If loListName.Includes(LCase(actLo.Name)) Then
                    Set actcsTab = CustomTable.Create(actLo)
                    Set impLo = impsh.ListObjects(actLo.Name)
                    Set impcsTab = CustomTable.Create(impLo)
                    actcsTab.Import impcsTab
                End If
            Next
            actsh.calculate
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
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set tradsh = wb.Worksheets("DesignerTranslation")
    Set desTrads = DesTranslation.Create(tradsh)

    tradsh.Range(RNGLANGCODE).Value = langId
    tradsh.calculate
    desTrads.TranslateDesigner mainsh

    'Update all the labels on the ribbon
    UpdateLabels

ExitLang:
    NotBusyApp
End Sub

'@Description("Callback for btnOpen onAction: Open another linelist file")
'@EntryPoint
Public Sub clickOpen(control As IRibbonControl)

    Dim io As IOSFiles
    Dim trads As IDesTranslation
    Set io = OSFiles.Create()

    io.LoadFile "*.xlsb"                         '
    If Not io.HasValidFile Then Exit Sub

    On Error GoTo ErrorManage
    NotBusyApp
    Application.Workbooks.Open fileName:=io.File(), ReadOnly:=False
    Exit Sub

ErrorManage:
    On Error Resume Next
    Set trads = DesTranslation.Create(ThisWorkbook.Worksheets(DESIGNERTRADSHEET))
    MsgBox trads.TranslationMsg("MSG_TitlePassWord"), vbCritical, _
    trads.TranslationMsg("MSG_PassWord")
    On Error GoTo 0
End Sub

'@Description("Callback for btnImpPass onAction: Import passwords from a worksheet")
'@EntryPoint
Public Sub clickImpPass(control As IRibbonControl)

    Dim io As IOSFiles
    Dim wb As Workbook
    Dim imppass As ILLPasswords
    Dim actpass As ILLPasswords

    Set io = OSFiles.Create()

    BusyApp
    io.LoadFile "*.xlsx"                         '
    If Not io.HasValidFile Then Exit Sub

    On Error GoTo ErrorManage
        
        Set wb = Workbooks.Open(fileName:=io.File(), ReadOnly:=False)
        Set imppass = LLPasswords.Create(wb.Worksheets(1))
        Set actpass = LLPasswords.Create(ThisWorkbook.Worksheets(PASSWORDSHEET))
        actpass.Import imppass
        wb.Close saveChanges:=False

ErrorManage:
    On Error Resume Next
    wb.Close saveChanges:=False
    NotBusyApp
    MsgBox "Done!"
    On Error GoTo 0
End Sub


'@EntryPoint
Public Sub clickImpStyle(control As IRibbonControl)

    Dim io As IOSFiles
    Dim wb As Workbook
    Dim currformatObj As ILLFormat
    
    Set io = OSFiles.Create()
    
    BusyApp
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile Then Exit Sub

    On Error GoTo ErrorManage

    Set wb = Workbooks.Open(fileName:=io.File(), ReadOnly:=False)
    Set currformatObj = LLFormat.Create(ThisWorkbook.Worksheets(FORMATSHEET))

    currformatObj.Import wb.Worksheets(1)
    wb.Close saveChanges:=False
    MsgBox "Done!"

ErrorManage:
    On Error Resume Next
    wb.Close saveChanges:=False
    NotBusyApp
    On Error GoTo 0
End Sub


'Callback for chkAlert getPressed
Public Sub initMainAlert(control As IRibbonControl, ByRef returnedVal)

    Dim mainsh As Worksheet
    
    Set mainsh = ThisWorkbook.Worksheets(DESIGNERMAINSHEET)
    
    If mainsh.Range("RNG_MainAlert").Value = "alert" Then
        returnedVal = False
    Else
        returnedVal = True
    End If
End Sub

'Callback for chkAlert onAction
Public Sub clickMainAlert(control As IRibbonControl, pressed As Boolean)

    Dim mainsh As Worksheet
    Set mainsh = ThisWorkbook.Worksheets(DESIGNERMAINSHEET)
    
    Select Case pressed
        Case True
        mainsh.Range("RNG_MainAlert").Value = "avoid alerts"
        Case False
        mainsh.Range("RNG_MainAlert").Value = "alert"
    End Select
End Sub


'Callback for chkInstruct getPressed
Public Sub initMainInstruct(control As IRibbonControl, ByRef returnedVal)

    Dim mainsh As Worksheet
    
    Set mainsh = ThisWorkbook.Worksheets(DESIGNERMAINSHEET)
    
    If mainsh.Range("RNG_MainInstruct").Value = "add" Then
        returnedVal = True
    Else
        returnedVal = False
    End If
End Sub

'Callback for chkAlert onAction
Public Sub clickMainInstruct(control As IRibbonControl, pressed As Boolean)

    Dim mainsh As Worksheet
    Set mainsh = ThisWorkbook.Worksheets(DESIGNERMAINSHEET)
    
    Select Case pressed
        Case True
        mainsh.Range("RNG_MainInstruct").Value = "add"
        Case False
        mainsh.Range("RNG_MainInstruct").Value = "not add"
    End Select
End Sub