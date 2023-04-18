Attribute VB_Name = "EventsDesignerRibbon"
Attribute VB_Description = "Events associated to the Ribbon Menu in the designer"
Option Explicit
Option Private Module

'@Folder("Designer Events")
'@ModuleDescription("Events associated to the Ribbon Menu in the designer")

'Designer Translation sheet name
Private Const DESIGNERTRADSHEET As String = "DesignerTranslation"
'Setup translation sheet name
Private Const SETUPTRADSHEET As String = "Translations"
'Linelist translation sheet name
Private Const LINELISTTRADSHEET As String = "LinelistTranslation"
'Designer main sheet name
Private Const DESIGNERMAINSHEET As String = "Main"
'Range for informations to user in the main sheet
Private Const RNGEDITION As String = "RNG_Edition"

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

    Set wb = ThisWorkbook
    Set mainsh = wb.Worksheets(DESIGNERMAINSHEET)

    mainsh.Range("RNG_PathDico").Value = vbNullString
    mainsh.Range("RNG_PathGeo").Value = vbNullString
    mainsh.Range("RNG_LLName").Value = vbNullString
    mainsh.Range("RNG_LLDir").Value = vbNullString
    mainsh.Range("RNG_Edition").Value = vbNullString
    mainsh.Range("RNG_Update").Value = vbNullString
    mainsh.Range("RNG_LangSetup").Value = vbNullString

    'Set all input ranges to while color
    DesignerMain.SetInputRangesToWhite
    mainsh.Range("RNG_Update").Interior.color = vbWhite
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
                Set actcsTab = CustomTable.Create(Lo)
                Set impLo = impsh.ListObjects(Lo.Name)
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
    Set io = OSFiles.Create()

    io.LoadFile "*.xlsb"                         '
    If Not io.HasValidFile Then Exit Sub

    On Error GoTo ErrorManage
    NotBusyApp
    Application.Workbooks.Open FileName:=io.File(), ReadOnly:=False
    Exit Sub

ErrorManage:
    MsgBox DesignerMain.TranslateDesMsg("MSG_TitlePassWord"), vbCritical, _
    DesignerMain.TranslateDesMsg("MSG_PassWord")
End Sub
