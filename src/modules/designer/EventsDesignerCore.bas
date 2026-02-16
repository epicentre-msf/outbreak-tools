Attribute VB_Name = "EventsDesignerCore"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, IDesignerPreparation, RibbonDev, OSFiles, IOSFiles, BetterArray, CustomTable, ICustomTable, Passwords, IPasswords, LLFormat, ILLFormat, ApplicationState, IApplicationState, DesignerTranslation, IDesignerTranslation, HiddenNames, IHiddenNames
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const SHEET_FORMAT As String = "__formatter"
Private Const DESTRADSSHEET As String = "DesignerTranslation"
Private Const MAINSHEET As String = "Main"
Private Const PROMPT_TITLE As String = "Designer"
Private Const TAG_FORMATTER_IMPORTED As String = "TAG_FORMATTER_IMPORTED"

Private trads As IDesignerTranslation
Private prep As IDesignerPreparation

'@section Ribbon lifecycle
'===============================================================================

'@Description("Return the translated label for a control; fallback to the control id.")
'@EntryPoint
Public Sub LangLabel(ByRef control As IRibbonControl, ByRef returnedVal)

    On Error GoTo Fallback
    EnsureTranslation
    If trads Is Nothing Then GoTo Fallback
    returnedVal = trads.TranslatedValue(control.Id)
    Exit Sub

Fallback:
    returnedVal = control.Id
End Sub

Private Sub EnsureTranslation()

    On Error Resume Next
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(DESTRADSSHEET)
    On Error GoTo 0

    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    If trads Is Nothing Then
        Set trads = DesignerTranslation.Create(sh)
    End If
    On Error GoTo 0
End Sub

Private Sub InvalidateRibbon()
    Dim ribbon As IRibbonUI
    Set ribbon = RibbonDev.ActualRibbon()
    If Not ribbon Is Nothing Then ribbon.Invalidate
End Sub


'@Description("Switch designer language and re-run translations.")
'@EntryPoint
Public Sub clickLangChange(ByRef control As IRibbonControl, ByRef langId As String, ByRef Index As Integer)
    Dim targetSheet As Worksheet
    Dim appScope As IApplicationState

    On Error GoTo Cleanup
    Set targetSheet = ThisWorkbook.Worksheets(MAINSHEET)
    Set appScope = ApplicationState.Create(Application)

    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False
    EnsureTranslation

    If trads Is Nothing Then GoTo Cleanup
    trads.TranslateDesigner targetSheet, langId
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
        On Error Resume Next
        Set targetSheet = targetBook.Worksheets(CStr(sheetName))
        Set sourceSheet = importBook.Worksheets(CStr(sheetName))
        On Error GoTo 0

        If (targetSheet Is Nothing) Or (sourceSheet Is Nothing) Then GoTo Cleanup

        For Each lo In targetSheet.ListObjects
            If tableNames.Includes(LCase$(lo.Name)) Then
                Set targetTable = CustomTable.Create(lo)
                Set sourceTable = CustomTable.Create(sourceSheet.ListObjects(lo.Name))
                targetTable.Import sourceTable
            End If
        Next lo
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

    Dim store As IHiddenNames
    Set store = HiddenNames.Create(ThisWorkbook)
    store.Add TAG_FORMATTER_IMPORTED, "Yes"

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
    returnedVal = ResolvePreparation().GetFlag("chkAlert", True)
End Sub

'@Description("Persist alert checkbox state.")
'@EntryPoint
Public Sub clickMainAlert(ByRef control As IRibbonControl, ByVal pressed As Boolean)
    ResolvePreparation().SetFlag "chkAlert", pressed
End Sub

'@Description("Initialise checkbox state for instructions from persisted hidden names.")
'@EntryPoint
Public Sub initMainInstruct(ByRef control As IRibbonControl, ByRef returnedVal)
    returnedVal = ResolvePreparation().GetFlag("chkInstruct", True)
End Sub

'@Description("Persist instruction checkbox state.")
'@EntryPoint
Public Sub clickMainInstruct(ByRef control As IRibbonControl, ByVal pressed As Boolean)
    ResolvePreparation().SetFlag "chkInstruct", pressed
End Sub


'@section Internal helpers
'===============================================================================

'@Description("Lazily resolve and cache the designer preparation helper bound to ThisWorkbook.")
Private Function ResolvePreparation() As IDesignerPreparation
    If prep Is Nothing Then
        Set prep = DesignerPreparation.Create(ThisWorkbook)
    End If
    Set ResolvePreparation = prep
End Function
