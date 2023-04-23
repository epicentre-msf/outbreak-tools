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
'Range for informations to user in the main sheet

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


'@Description("Callback for getLabel (Depending on the language)")
'@EntryPoint
Public Sub LangLabel(control As IRibbonControl, ByRef returnedVal)
Attribute LangLabel.VB_Description = "Callback for getLabel (Depending on the language)"
End Sub

'@Description("Callback for btnDelGeo onAction: Delete the geobase")
'@EntryPoint
Public Sub clickDelGeo(control As IRibbonControl)
Attribute clickDelGeo.VB_Description = "Callback for btnDelGeo onAction: Delete the geobase"
End Sub

'@Description("Callback for btnClear onAction": Clear the entries)
'@EntryPoint
Public Sub clickClearEnt(control As IRibbonControl)

    Dim wb As Workbook
    Dim mainsh As Worksheet
    Dim mainobj As IMain

    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)
    Set mainobj = Main.Create(mainsh)
    mainobj.ClearInputRanges clearValues := True
End Sub

'@Description("Callback for btnTransAdd onAction: Import Linelist translations")
'@EntryPoint
Public Sub clickImpTrans(control As IRibbonControl)
Attribute clickImpTrans.VB_Description = "Callback for btnTransAdd onAction: Import Linelist translations"

    Const TRADSHEETNAME As String = "LinelistTranslation"

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

    Set wb = ThisWorkbook
    Set actsh = wb.Worksheets(TRADSHEETNAME)

    'Import the translations for
    Set io = OSFiles.Create()
    Set loListName = New BetterArray

    io.LoadFile "*.xlsb"
    If io.HasValidFile() Then
        BusyApp
        Set impwb = Workbooks.Open(io.File())
        On Error GoTo ExitTrads
        Set impsh = impwb.Worksheets(TRADSHEETNAME)
        For Each actLo In actsh.ListObjects
            If (actLo.Name = "T_TradLLShapes") Or _
               (actLo.Name = "T_TradLLMsg") Or _
               (actLo.Name = "T_TradLLForms") Then
                Set actcsTab = CustomTable.Create(actLo)
                Set impLo = impsh.ListObjects(actLo.Name)
                Set impcsTab = CustomTable.Create(impLo)
                actcsTab.Import impcsTab
            End If
        Next
        On Error GoTo 0
    End If
ExitTrads:
    On Error Resume Next
    impwb.Close saveChanges:=False
    NotBusyApp
    On Error GoTo 0
End Sub

'@Description("Callback for langDrop onAction: Change the language of the designer")
'@EntryPoint
Public Sub clickLangChange(control As IRibbonControl, id As String, Index As Integer)
Attribute clickLangChange.VB_Description = "Callback for langDrop onAction: Change the language of the designer"
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
