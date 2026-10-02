VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "WithEvents wrapper for export command buttons"

'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, ProcedureNotUsed, HungarianNotation
'@PredeclaredId
'@Folder("DataIO")
'@ModuleDescription("WithEvents wrapper for export command buttons")

Option Explicit

'@class ExportButton
'@description
'WithEvents wrapper that listens to an MSForms.CommandButton on a linelist
'export sheet and performs a custom export when clicked. The factory wires
'the button, an optional filter checkbox, and translation/workbook
'references. On click, RunExport selects a folder, optionally confirms
'filter usage, syncs the filtered companions through FilteredData, then
'delegates to LLExporter.ExportCustom. A sheet the sync skipped aborts the
'export through the click handler's error path. The factory returns the
'concrete class because the caller holds a reference to keep the WithEvents
'binding alive. That caller is the F_Export code-behind, which keeps one
'instance per button in a module-level collection for as long as the form
'carries the buttons.
'@depends LLExporter, FilteredData, ApplicationState, OSFiles, TranslationObject, BetterArray


'@section Constants
'===============================================================================

Private Const CLASS_NAME As String = "ExportButton"
Private Const NAMETAG As String = "CMDExport"


'@section State
'===============================================================================

Private WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Private chkFilt As MSForms.CheckBox

Private Type TExportButton
    sourceWkb As Workbook
    trads As TranslationObject
    sealed As Boolean
    running As Boolean
End Type

Private this As TExportButton


'@section Factory
'===============================================================================

'@label:create
'@sub-title Create an export button wrapper
'@details
'Factory method that validates the workbook, translations, and button are
'not Nothing, then wires the WithEvents binding and optional filter
'checkbox. Called on the predeclared instance. Returns the concrete class
'because the caller holds a reference to keep the WithEvents binding alive.
'@param sourceWkb Workbook. The linelist workbook.
'@param trads TranslationObject. Translations for user-facing messages.
'@param expBtn MSForms.CommandButton. The button to listen to.
'@param expChk Optional MSForms.CheckBox. Filter checkbox, or Nothing. Defaults to Nothing.
'@return ExportButton. A wired instance ready to handle click events.
'@throws ProjectError.ObjectNotInitialized When sourceWkb, trads, or expBtn is Nothing.
'@export
Public Function Create(ByVal sourceWkb As Workbook, _
                       ByVal trads As TranslationObject, _
                       ByVal expBtn As MSForms.CommandButton, _
                       Optional ByVal expChk As MSForms.CheckBox = Nothing) As ExportButton

    If sourceWkb Is Nothing Then
        ThrowError ObjectNotInitialized, "Workbook required"
    End If
    If trads Is Nothing Then
        ThrowError ObjectNotInitialized, "Translations required"
    End If
    If expBtn Is Nothing Then
        ThrowError ObjectNotInitialized, "Button required"
    End If

    Dim newButton As ExportButton
    Set newButton = New ExportButton
    Set newButton.SourceWorkbook = sourceWkb
    Set newButton.Translations = trads
    Set newButton.PushButton = expBtn
    If Not expChk Is Nothing Then Set newButton.FilterCheckBox = expChk
    newButton.Seal
    Set Create = newButton
End Function


'@section Public Accessors
'===============================================================================

'@label:source-workbook-set
'@prop-title Assign the source linelist workbook
'@param wb Workbook. The workbook to store.
Public Property Set SourceWorkbook(ByVal wb As Workbook)
    GuardNotSealed "SourceWorkbook"
    Set this.sourceWkb = wb
End Property

'@label:translations-set
'@prop-title Assign the translations object
'@param trads TranslationObject. The translations to store.
Public Property Set Translations(ByVal trads As TranslationObject)
    GuardNotSealed "Translations"
    Set this.trads = trads
End Property

'@label:push-button-set
'@prop-title Wire the WithEvents button binding
'@param expBtn MSForms.CommandButton. The button to listen to.
Public Property Set PushButton(ByVal expBtn As MSForms.CommandButton)
    GuardNotSealed "PushButton"
    Set Btn = expBtn
End Property

'@label:filter-checkbox-set
'@prop-title Assign the optional filter checkbox
'@param expChk MSForms.CheckBox. The checkbox to bind.
Public Property Set FilterCheckBox(ByVal expChk As MSForms.CheckBox)
    GuardNotSealed "FilterCheckBox"
    Set chkFilt = expChk
End Property

'@label:seal
'@sub-title Prevent further changes to setup-only properties
Public Sub Seal()
    this.sealed = True
End Sub

'@label:guard-not-sealed
'@sub-title Raise when a setup-only property is written after sealing
'@param propName String. The name of the property being guarded.
Private Sub GuardNotSealed(ByVal propName As String)
    If this.sealed Then _
        ThrowError ProjectError.SomethingWentWrong, propName & " is set at creation only"
End Sub

'@label:export-number
'@prop-title Export number derived from button name
'@details
'Reads the button Name property, takes the NAMETAG prefix ("CMDExport") off
'the front and reads the digits that follow. "CMDExport1" answers 1 and
'"CMDExport3" answers 3.
'
'The property answers 0 for every name it cannot place: no button bound, a
'name that does not lead with the tag, a remainder with no digits in it, and
'a number outside the Long range. The read runs on the click path, so it
'raises nothing and RunExport refuses the 0.
'@return Long. The 1-based export number, or 0 when the name carries none.
'@export
Public Property Get ExportNumber() As Long

    Dim buttonName As String
    Dim numericPart As Double

    If Btn Is Nothing Then Exit Property

    buttonName = Btn.Name
    If StrComp(Left$(buttonName, Len(NAMETAG)), NAMETAG, vbTextCompare) <> 0 Then _
        Exit Property

    'Val reads the leading digits and answers 0 for a remainder with none, so
    'it never raises the way CLng of a word does. Int drops any fraction, so
    'the answer is the number that was written.
    numericPart = Int(Val(Mid$(buttonName, Len(NAMETAG) + 1)))
    If numericPart < 1 Or numericPart > 2147483647# Then Exit Property

    ExportNumber = CLng(numericPart)
End Property

'@label:use-filter
'@prop-title Whether the filter checkbox is checked
'@details
'Returns the current value of the filter checkbox. Returns False when
'no checkbox was bound.
'
'The box is read only. Every export button on the form reads the same one,
'and two exports never run at the same time, so the box holds the user's
'standing wish and the confirmation is answered per export.
'@return Boolean. True when the filter checkbox is checked.
'@export
Public Property Get UseFilter() As Boolean
    If Not chkFilt Is Nothing Then UseFilter = chkFilt.Value
End Property


'@section Event Handler
'===============================================================================

'@label:btn-click
'@sub-title Handle button click event
'@details
'A click that arrives while this button is already exporting is dropped.
'Screen updating is off for most of an export, so the form does not repaint
'and a user who reads that as nothing happening clicks again. The queued
'click runs when the first export returns, and it would ask for a folder a
'second time and write a second file.
Private Sub Btn_Click()

    If this.running Then Exit Sub

    this.running = True
    RunExport
    this.running = False
End Sub

'@label:run-export
'@sub-title Perform a custom export
'@details
'Asks the user everything it needs first: the folder through OSFiles, then
'the filter confirmation. Both questions are answered before the busy state
'goes on, because a box drawn under suppressed screen updating can end up
'behind the workbook window on Mac Excel. The busy state then covers the
'work: the filtered companions are synced and LLExporter.ExportCustom writes
'the file. The state is restored before the user is told where the file went.
'On error, displays a translated error message and closes any orphaned
'output workbook.
'@throws ProjectError.SomethingWentWrong When the button name carries no
'export number.
Private Sub RunExport()

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim folderPath As String
    Dim filePath As String
    Dim useFilt As Boolean
    Dim failureDetail As String
    Dim runNumber As Long

    On Error GoTo ErrHand

    ' The export this button runs
    runNumber = ExportNumber
    If runNumber < 1 Then _
        ThrowError SomethingWentWrong, "The button name carries no export number"

    ' Select export folder
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    ' Handle filter confirmation
    useFilt = UseFilter
    If useFilt Then
        If MsgBox(TranslatedValue("MSG_AskFilter"), _
                  vbYesNo + vbQuestion, _
                  TranslatedValue("MSG_ThereIsFilter")) = vbNo Then
            'The refusal stays with this export. The box is shared by every
            'button of the form, and unticking it here would answer for the
            'next export the user runs.
            useFilt = False
        End If
    End If

    ' Apply busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    ' Sync the filtered companions the export reads
    If useFilt Then SyncFilteredData

    ' Export
    Set exporter = LLExporter.Create(this.sourceWkb)
    filePath = exporter.ExportCustom(runNumber, folderPath, useFilt)

    appState.Restore

    ReportSavedFile filePath, exporter.LastExportPassword
    Exit Sub

ErrHand:
    'Any On Error statement resets the Err object, so the detail is read
    'before the handler shields itself.
    failureDetail = Err.Description
    On Error Resume Next
    MsgBox TranslatedValue("MSG_ErrHandExport") & ": " & failureDetail, _
           vbOKOnly + vbCritical, TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub

'@section Private Helpers
'===============================================================================

'@label:translated-value
'@sub-title Get a translated message value
'@param msgCode String. The message code to translate.
'@return String. The translated message text.
Private Function TranslatedValue(ByVal msgCode As String) As String
    TranslatedValue = this.trads.TranslatedValue(msgCode)
End Function

'@label:report-saved-file
'@sub-title Tell the user where the file went and what opens it
'@details
'A password-protected export is encrypted with the linelist private key, and
'that key sits in a very hidden worksheet. A user who is not told it here has
'no way back to the file, and neither has the colleague they send it to. The
'key is one value for the whole workbook, so every password-protected export
'from this linelist shares it and generating a new key changes it for all of
'them.
'
'Nothing is shown when the export answered no path, which is what a mode that
'stopped early returns.
'@param filePath String. The saved file path.
'@param password String. The password applied to the file, or vbNullString.
Private Sub ReportSavedFile(ByVal filePath As String, ByVal password As String)

    Dim message As String

    If LenB(filePath) = 0 Then Exit Sub

    message = filePath & vbNewLine & vbNewLine

    If LenB(password) > 0 Then
        message = message & TranslatedValue("MSG_Password") & " " & password
    Else
        message = message & TranslatedValue("MSG_NoPassword")
    End If

    MsgBox message, vbOKOnly + vbInformation, TranslatedValue("MSG_FileSaved")
End Sub

'@label:sync-filtered-data
'@sub-title Sync the filtered companions and abort on a skipped sheet
'@details
'Builds a FilteredData over the source workbook and runs the sync. A sheet
'the sync skipped means the export would read a companion that no longer
'matches its table, so the failures are joined into one message and raised.
'The raise lands in RunExport's handler, which shows the message, closes
'any orphaned output workbook and restores the application state.
'@throws ProjectError.SomethingWentWrong When the sync skipped a sheet.
Private Sub SyncFilteredData()

    Dim filtered As FilteredData
    Dim failures As BetterArray
    Dim failureText As String
    Dim counter As Long

    Set filtered = FilteredData.Create(this.sourceWkb)
    filtered.Sync

    Set failures = filtered.FailedSheets
    If failures.Length = 0 Then Exit Sub

    For counter = failures.LowerBound To failures.UpperBound
        failureText = failureText & vbNewLine & failures.Item(counter)
    Next counter

    'The MSG_ErrUpdate row is still owed to the translation workbooks, and
    'TranslateTag answers the tag itself on a miss, so the box shows the raw
    'key until the row lands.
    ThrowError SomethingWentWrong, TranslatedValue("MSG_ErrUpdate") & failureText
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. A ProjectError enum value (e.g. ObjectNotInitialized).
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
