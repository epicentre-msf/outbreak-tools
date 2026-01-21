Attribute VB_Name = "EventsDesignerRibbon"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the designer workbook.")
'@depends DesignerEntry, IDesignerEntry, DesignerPreparation, IDesignerPreparation, RibbonDev, LLGeo, ILLGeo, OSFiles, IOSFiles, BetterArray, CustomTable, ICustomTable, Passwords, IPasswords, LLFormat, ApplicationState, IApplicationState, DesTranslation, IDesTranslation
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const SHEET_GEO As String = "Geo"
Private Const SHEET_FORMAT As String = "LinelistStyle"
Private Const PROMPT_TITLE As String = "Designer"

Private gRibbon As IRibbonUI


'@section Ribbon lifecycle
'===============================================================================
'@Description("Capture the ribbon instance when the UI loads.")
'@EntryPoint
Public Sub ribbonLoaded(ByRef ribbon As IRibbonUI)
    Static ribbonRegistered As Boolean

    If ribbonRegistered Then Exit Sub

    Set gRibbon = ribbon
    ribbonRegistered = True
End Sub

'@Description("Return the captured ribbon instance when available.")
Public Function ActualRibbon() As IRibbonUI
    If Not gRibbon Is Nothing Then
        Set ActualRibbon = gRibbon
    Else
        Set ActualRibbon = RibbonDev.ActualRibbon
    End If
End Function

'@Description("Return the translated label for a control; fallback to the control id.")
'@EntryPoint
Public Sub LangLabel(ByRef control As IRibbonControl, ByRef returnedVal)
    Dim translations As IDesTranslation

    On Error GoTo Fallback
    Set translations = ResolveTranslations()
    returnedVal = translations.TranslationMsg(control.Id)
    Exit Sub

Fallback:
    returnedVal = control.Id
End Sub


'@section Manage group callbacks
'===============================================================================
'@Description("Delete geobase content from the Geo worksheet.")
'@EntryPoint
Public Sub clickDelGeo(ByRef control As IRibbonControl)
    Dim geoSheet As Worksheet
    Dim geo As ILLGeo
    Dim appScope As IApplicationState

    On Error GoTo Cleanup

    Set geoSheet = ThisWorkbook.Worksheets(SHEET_GEO)
    Set geo = LLGeo.Create(geoSheet)
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    geo.Clear

Cleanup:
    If Not appScope Is Nothing Then appScope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickDelGeo: "; Err.Number; Err.Description
        MsgBox "Unable to clear the geobase: " & Err.Description, vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub

'@Description("Clear designer entry ranges on the active worksheet.")
'@EntryPoint
Public Sub clickClearEnt(ByRef control As IRibbonControl)
    Dim targetSheet As Worksheet
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

    If Not TypeName(ActiveSheet) = "Worksheet" Then Exit Sub
    Set targetSheet = ActiveSheet

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set entry = DesignerEntry.Create(targetSheet)
    entry.Clear

Cleanup:
    If Not appScope Is Nothing Then appScope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickClearEnt: "; Err.Number; Err.Description
        MsgBox "Unable to clear entries: " & Err.Description, vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub

'@Description("Switch designer language and re-run translations.")
'@EntryPoint
Public Sub clickLangChange(ByRef control As IRibbonControl, ByRef langId As String, ByRef Index As Integer)
    Dim targetSheet As Worksheet
    Dim entry As IDesignerEntry
    Dim appScope As IApplicationState

    If Not TypeName(ActiveSheet) = "Worksheet" Then Exit Sub
    Set targetSheet = ActiveSheet

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set entry = DesignerEntry.Create(targetSheet)
    entry.Translate langId
    InvalidateRibbon

Cleanup:
    If Not appScope Is Nothing Then appScope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickLangChange: "; Err.Number; Err.Description
        MsgBox "Unable to change language: " & Err.Description, vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub


'@section Import group callbacks
'===============================================================================
'@Description("Import translations tables from an external workbook.")
'@EntryPoint
Public Sub clickImpTrans(ByRef control As IRibbonControl)
    Dim io As IOSFiles
    Dim importBook As Workbook
    Dim targetBook As Workbook
    Dim sheetNames As BetterArray
    Dim tableNames As BetterArray
    Dim sheetName As Variant
    Dim idx As Long
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lo As ListObject
    Dim targetTable As ICustomTable
    Dim sourceTable As ICustomTable
    Dim appScope As IApplicationState

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set targetBook = ThisWorkbook
    Set sheetNames = New BetterArray
    sheetNames.Push "LinelistTranslation", "DesignerTranslation"

    Set tableNames = New BetterArray
    tableNames.Push "t_tradllshapes", "t_tradllmsg", "t_tradllforms", "t_tradllribbon", _
                    "t_tradmsg", "t_tradrange", "t_tradshape"

    Set importBook = Workbooks.Open(io.File())

    For idx = sheetNames.LowerBound To sheetNames.UpperBound
        sheetName = sheetNames.Item(idx)
        Set targetSheet = targetBook.Worksheets(CStr(sheetName))
        Set sourceSheet = importBook.Worksheets(CStr(sheetName))

        For Each lo In targetSheet.ListObjects
            If tableNames.Includes(LCase$(lo.Name)) Then
                Set targetTable = CustomTable.Create(lo)
                Set sourceTable = CustomTable.Create(sourceSheet.ListObjects(lo.Name))
                targetTable.Import sourceTable
            End If
        Next lo

        targetSheet.Calculate
    Next idx

    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE

Cleanup:
    If Not importBook Is Nothing Then importBook.Close saveChanges:=False
    If Not appScope Is Nothing Then appScope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickImpTrans: "; Err.Number; Err.Description
        MsgBox "Unable to import translations: " & Err.Description, vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub

'@Description("Import passwords from an external workbook.")
'@EntryPoint
Public Sub clickImpPass(ByRef control As IRibbonControl)
    Dim io As IOSFiles
    Dim importBook As Workbook
    Dim importer As IPasswords
    Dim target As IPasswords
    Dim appScope As IApplicationState

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=False)
    Set importer = Passwords.Create(importBook.Worksheets(1))
    Set target = Passwords.Create(ThisWorkbook.Worksheets("__pass"))
    target.Import importer

    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE

Cleanup:
    If Not importBook Is Nothing Then importBook.Close saveChanges:=False
    If Not appScope Is Nothing Then appScope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickImpPass: "; Err.Number; Err.Description
        MsgBox "Unable to import passwords: " & Err.Description, vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub

'@Description("Import linelist format from a workbook.")
'@EntryPoint
Public Sub clickImpStyle(ByRef control As IRibbonControl)
    Dim io As IOSFiles
    Dim importBook As Workbook
    Dim formatManager As ILLFormat
    Dim appScope As IApplicationState

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=False)
    Set formatManager = LLFormat.Create(ThisWorkbook.Worksheets(SHEET_FORMAT))
    formatManager.Import importBook.Worksheets(1)

    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE

Cleanup:
    If Not importBook Is Nothing Then importBook.Close saveChanges:=False
    If Not appScope Is Nothing Then appScope.Restore
    If Err.Number <> 0 Then
        Debug.Print "clickImpStyle: "; Err.Number; Err.Description
        MsgBox "Unable to import styles: " & Err.Description, vbExclamation + vbOKOnly, PROMPT_TITLE
        Err.Clear
    End If
End Sub


'@section Advanced group callbacks
'===============================================================================
'@Description("Open a linelist workbook selected by the user.")
'@EntryPoint
Public Sub clickOpen(ByRef control As IRibbonControl)
    Dim io As IOSFiles

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsb"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo ErrorManage
    Workbooks.Open fileName:=io.File(), ReadOnly:=False
    Exit Sub

ErrorManage:
    MsgBox "Unable to open workbook: " & Err.Description, vbCritical + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@Description("Initialise checkbox state for alerts from persisted hidden names.")
'@EntryPoint
Public Sub initMainAlert(ByRef control As IRibbonControl, ByRef returnedVal)
    returnedVal = ResolvePreparation().GetFlag("chkAlert", False)
End Sub

'@Description("Persist alert checkbox state.")
'@EntryPoint
Public Sub clickMainAlert(ByRef control As IRibbonControl, ByVal pressed As Boolean)
    ResolvePreparation().SetFlag "chkAlert", pressed
End Sub

'@Description("Initialise checkbox state for instructions from persisted hidden names.")
'@EntryPoint
Public Sub initMainInstruct(ByRef control As IRibbonControl, ByRef returnedVal)
    returnedVal = ResolvePreparation().GetFlag("chkInstruct", False)
End Sub

'@Description("Persist instruction checkbox state.")
'@EntryPoint
Public Sub clickMainInstruct(ByRef control As IRibbonControl, ByVal pressed As Boolean)
    ResolvePreparation().SetFlag "chkInstruct", pressed
End Sub


'@section Multi group callbacks (to be implemented later)
'===============================================================================
Public Sub clickFolderMulti(ByRef control As IRibbonControl)
    NotifyPlannedWork
End Sub

Public Sub clickDupMulti(ByRef control As IRibbonControl)
    NotifyPlannedWork
End Sub

Public Sub clickAddRowsMulti(ByRef control As IRibbonControl)
    NotifyPlannedWork
End Sub

Public Sub clickResizeMulti(ByRef control As IRibbonControl)
    NotifyPlannedWork
End Sub

Public Sub clickImpMulti(ByRef control As IRibbonControl)
    NotifyPlannedWork
End Sub

Public Sub clickExportMulti(ByRef control As IRibbonControl)
    NotifyPlannedWork
End Sub


'@section Dev group callbacks
'===============================================================================
'@Description("Initialise designer preparation using development helpers.")
'@EntryPoint
Public Sub clickDevInitialize(ByRef control As IRibbonControl)
    Dim prep As IDesignerPreparation

    Set prep = ResolvePreparation()
    prep.Prepare RibbonDev.EnsureDevelopment()
    MsgBox "Done!", vbInformation + vbOKOnly, PROMPT_TITLE
End Sub

Public Sub clickDevFolder(ByRef control As IRibbonControl)
    RibbonDev.clickDevFolder control
End Sub

Public Sub clickDevImport(ByRef control As IRibbonControl)
    RibbonDev.clickDevImport control
End Sub

Public Sub clickDevExport(ByRef control As IRibbonControl)
    RibbonDev.clickDevExport control
End Sub

Public Sub clickDevVBE(ByRef control As IRibbonControl)
    RibbonDev.clickDevVBE control
End Sub

Public Sub clickDevDeploy(ByRef control As IRibbonControl)
    RibbonDev.clickDevDeploy control
End Sub

Public Sub clickDevAddRows(ByRef control As IRibbonControl)
    RibbonDev.clickDevAddRows control
End Sub

Public Sub clicDevResize(ByRef control As IRibbonControl)
    RibbonDev.clicDevResize control
End Sub

Public Sub clicDevAddFormTable(ByRef control As IRibbonControl)
    RibbonDev.clicDevAddFormTable control
End Sub

Public Sub clickDevAddClassTable(ByRef control As IRibbonControl)
    RibbonDev.clickDevAddClassTable control
End Sub

Public Sub clickDevAddFormTable(ByRef control As IRibbonControl)
    RibbonDev.clickDevAddFormTable control
End Sub

Public Sub clickDevAddModulesTable(ByRef control As IRibbonControl)
    RibbonDev.clickDevAddModulesTable control
End Sub


'@section Helpers
'===============================================================================
Private Sub NotifyPlannedWork()
    MsgBox "This ribbon action will be implemented in a future update.", vbInformation + vbOKOnly, PROMPT_TITLE
End Sub

Private Sub InvalidateRibbon()
    Dim ribbon As IRibbonUI
    Set ribbon = ActualRibbon()
    If Not ribbon Is Nothing Then ribbon.Invalidate
End Sub

Private Function ResolvePreparation() As IDesignerPreparation
    Set ResolvePreparation = DesignerPreparation.Create(ThisWorkbook)
End Function

Private Function ResolveTranslations() As IDesTranslation
    Set ResolveTranslations = DesTranslation.Create(ThisWorkbook.Worksheets("DesignerTranslation"))
End Function
