Attribute VB_Name = "EventsDesignerCore"
Option Explicit

'@Folder("Designer")
'@ModuleDescription("Ribbon callbacks for the designer workbook.")
'@depends DesignerPreparation, DesignerEntry, RibbonDev, OSFiles, Passwords, LLFormat, ApplicationState, DesignerTranslation
'@IgnoreModule UnrecognizedAnnotation, ParameterNotUsed, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName

Private Const SHEET_FORMAT As String = "__formatter"
Private Const SHEET_PASS As String = "__pass"
Private Const DESTRADSSHEET As String = "DesignerTranslation"
Private Const MAINSHEET As String = "Main"
Private Const PROMPT_TITLE As String = "Designer"

'The designer holds one entry-and-translator pair for the whole session.
'Every callback used to build its own DesignerEntry, and the entry lazily
'built a DesignerTranslation, which reads four translation tables -- per
'press. The pair is dropped when the language changes and when an import
'rewrites the tables.
'
'triedTranslation is the flag the linelist event surface needed for the same
'reason: a build that fails leaves the field Nothing, and the guard "the
'field is Nothing" then reads as "never tried", so every press rebuilt and
're-failed the same object. The flag says the attempt was made.
Private trads As DesignerTranslation
Private prep As DesignerPreparation
Private entry As DesignerEntry
Private triedTranslation As Boolean

'@section Shared designer services
'===============================================================================

'@Description("The one DesignerEntry over the Main worksheet, with the shared translator.")
Public Function EntryManager() As DesignerEntry
    If entry Is Nothing Then
        Set entry = DesignerEntry.Create(ThisWorkbook.Worksheets(MAINSHEET))

        'The entry resolves a translator of its own on first use. Handing it
        'the held one keeps the pair to a single DesignerTranslation.
        EnsureTranslation
        If Not trads Is Nothing Then entry.UseTranslator trads
    End If

    Set EntryManager = entry
End Function

'@Description("Drop the held entry and translator so the next caller builds them again.")
Public Sub ResetDesignerCaches()
    Set entry = Nothing
    Set trads = Nothing
    triedTranslation = False
End Sub


'@section Ribbon lifecycle
'===============================================================================

'@Description("Return the translated label for a control; fallback to the control id.")
'@EntryPoint
Public Sub LangLabel(ByRef ribbonControl As IRibbonControl, ByRef returnedVal)

    On Error GoTo Fallback
    EnsureTranslation
    If trads Is Nothing Then GoTo Fallback
    returnedVal = trads.TranslatedValue(ribbonControl.Id)
    Exit Sub

Fallback:
    returnedVal = ribbonControl.Id
End Sub

Private Sub EnsureTranslation()
    Dim sh As Worksheet

    If triedTranslation Then Exit Sub
    triedTranslation = True

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(DESTRADSSHEET)
    On Error GoTo 0

    If sh Is Nothing Then Exit Sub

    On Error Resume Next
    Set trads = DesignerTranslation.Create(sh)
    On Error GoTo 0
End Sub

Private Sub InvalidateRibbon()
    Dim ribbon As IRibbonUI
    Set ribbon = RibbonDev.ActualRibbon()
    If Not ribbon Is Nothing Then ribbon.Invalidate
End Sub


'@Description("Switch designer language and re-run translations.")
'@EntryPoint
Public Sub clickLangChange(ByRef ribbonControl As IRibbonControl, ByRef langId As String, ByRef Index As Integer)
    Dim targetSheet As Worksheet
    Dim appScope As ApplicationState

    On Error GoTo Cleanup
    Set targetSheet = ThisWorkbook.Worksheets(MAINSHEET)
    Set appScope = ApplicationState.Create(Application)

    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False
    EnsureTranslation

    If trads Is Nothing Then GoTo Cleanup
    trads.TranslateDesigner targetSheet, langId
    InvalidateRibbon

    'The designer speaks another language now, so the pair is dropped and the
    'next callback reads the rows of the new language.
    ResetDesignerCaches

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
Public Sub clickImpTrans(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles
    Dim importBook As Workbook
    Dim appScope As ApplicationState

    'The picker runs before the busy state, because the dialog needs the UI
    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    'The import body lives on DesignerPreparation, which the preparation
    'sequence uses too. The workbook is opened here and closed here, so the
    'class reads an open book and leaves it alone.
    Set importBook = Workbooks.Open(io.File(), ReadOnly:=True)
    ResolvePreparation().ImportTranslations importBook

    'The translation tables were rewritten under the held translator.
    ResetDesignerCaches

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
Public Sub clickImpPass(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles
    Dim importBook As Workbook
    Dim importer As Passwords
    Dim target As Passwords
    Dim appScope As ApplicationState

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=False)
    Set importer = Passwords.Create(importBook.Worksheets(1))
    Set target = Passwords.Create(ThisWorkbook.Worksheets(SHEET_PASS))
    target.ImportFrom importer

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
Public Sub clickImpStyle(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles
    Dim importBook As Workbook
    Dim formatManager As LLFormat
    Dim appScope As ApplicationState

    Set io = OSFiles.Create()
    io.LoadFile "*.xlsx"
    If Not io.HasValidFile() Then Exit Sub

    On Error GoTo Cleanup
    Set appScope = ApplicationState.Create(Application)
    appScope.ApplyBusyState suppressEvents:=True, calculateOnSave:=False

    Set importBook = Workbooks.Open(io.File(), ReadOnly:=False)
    Set formatManager = LLFormat.Create(ThisWorkbook.Worksheets(SHEET_FORMAT))
    formatManager.Import importBook.Worksheets(1)

    'The designer now holds the live formatter. Loading a setup file clears
    'this flag again, so the styles imported here belong to the setup that is
    'loaded now.
    Dim designerPrep As DesignerPreparation
    Set designerPrep = ResolvePreparation()
    designerPrep.FormatterImported = True

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
Public Sub clickOpen(ByRef ribbonControl As IRibbonControl)
    Dim io As OSFiles

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
Public Sub initMainAlert(ByRef ribbonControl As IRibbonControl, ByRef returnedVal)

    'getPressed fires at ribbon load; a raise here drops the whole tab,
    'so any failure falls back to the checked default.
    On Error GoTo Fallback
    returnedVal = ResolvePreparation().GetFlag("chkAlert", True)
    Exit Sub

Fallback:
    returnedVal = True
End Sub

'@Description("Persist alert checkbox state.")
'@EntryPoint
Public Sub clickMainAlert(ByRef ribbonControl As IRibbonControl, ByVal pressed As Boolean)

    On Error GoTo Handler
    ResolvePreparation().SetFlag "chkAlert", pressed
    Exit Sub

Handler:
    Debug.Print "clickMainAlert: "; Err.Number; Err.Description
    MsgBox "Unable to save the alert setting: " & Err.Description, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub

'@Description("Initialise checkbox state for instructions from persisted hidden names.")
'@EntryPoint
Public Sub initMainInstruct(ByRef ribbonControl As IRibbonControl, ByRef returnedVal)

    'getPressed fires at ribbon load; a raise here drops the whole tab,
    'so any failure falls back to the checked default.
    On Error GoTo Fallback
    returnedVal = ResolvePreparation().GetFlag("chkInstruct", True)
    Exit Sub

Fallback:
    returnedVal = True
End Sub

'@Description("Persist instruction checkbox state.")
'@EntryPoint
Public Sub clickMainInstruct(ByRef ribbonControl As IRibbonControl, ByVal pressed As Boolean)

    On Error GoTo Handler
    ResolvePreparation().SetFlag "chkInstruct", pressed
    Exit Sub

Handler:
    Debug.Print "clickMainInstruct: "; Err.Number; Err.Description
    MsgBox "Unable to save the instruction setting: " & Err.Description, _
           vbExclamation + vbOKOnly, PROMPT_TITLE
    Err.Clear
End Sub


'@section Internal helpers
'===============================================================================

'@Description("Lazily resolve and cache the designer preparation helper bound to ThisWorkbook.")
Private Function ResolvePreparation() As DesignerPreparation
    If prep Is Nothing Then
        Set prep = DesignerPreparation.Create(ThisWorkbook)
    End If
    Set ResolvePreparation = prep
End Function
