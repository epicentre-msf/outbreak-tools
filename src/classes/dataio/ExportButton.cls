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
'filter usage, then delegates to LLExporter.ExportCustom. The factory
'returns the concrete class because callers hold a
'reference to keep the WithEvents binding alive.
'@depends LLExporter, ApplicationState, OSFiles


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
'(not an interface) because callers hold a reference to keep the
'WithEvents binding alive.
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
'Parses the button Name property by removing the NAMETAG prefix
'("CMDExport") and converting the remainder to Long. For example,
'"CMDExport1" yields 1, "CMDExport3" yields 3.
'@return Long. The 1-based export number.
'@export
Public Property Get ExportNumber() As Long
    ExportNumber = CLng(Replace(Btn.Name, NAMETAG, vbNullString))
End Property

'@label:use-filter
'@prop-title Whether the filter checkbox is checked
'@details
'Returns the current value of the filter checkbox. Returns False when
'no checkbox was bound.
'@return Boolean. True when the filter checkbox is checked.
'@export
Public Property Get UseFilter() As Boolean
    If Not chkFilt Is Nothing Then UseFilter = chkFilt.Value
End Property

'@label:use-filter-set
'@prop-title Set the filter checkbox value
'@param value Boolean. True to check, False to uncheck.
Public Property Let UseFilter(ByVal value As Boolean)
    If Not chkFilt Is Nothing Then chkFilt.Value = value
End Property


'@section Event Handler
'===============================================================================

'@label:btn-click
'@sub-title Handle button click event
Private Sub Btn_Click()
    RunExport
End Sub

'@label:run-export
'@sub-title Perform a custom export
'@details
'Prompts the user to select a folder via OSFiles, applies busy state,
'optionally confirms filter usage via MsgBox, then delegates to
'LLExporter.ExportCustom and tells the user where the file went. On error,
'displays a translated error message and closes any orphaned output workbook.
Private Sub RunExport()

    Dim exporter As LLExporter
    Dim appState As ApplicationState
    Dim io As OSFiles
    Dim folderPath As String
    Dim filePath As String
    Dim useFilt As Boolean

    On Error GoTo ErrHand

    ' Select export folder
    Set io = OSFiles.Create()
    io.LoadFolder
    If Not io.HasValidFolder Then Exit Sub
    folderPath = io.Folder()

    ' Apply busy state
    Set appState = ApplicationState.Create()
    appState.ApplyBusyState suppressEvents:=True, calculateOnSave:=False, _
                            busyCursor:=xlWait, blockSecurity:=False

    ' Handle filter confirmation
    useFilt = UseFilter
    If useFilt Then
        If MsgBox(TranslatedValue("MSG_AskFilter"), _
                  vbYesNo + vbQuestion, _
                  TranslatedValue("MSG_ThereIsFilter")) = vbYes Then
            UpdateFilteredData
        Else
            UseFilter = False
            useFilt = False
        End If
    End If

    ' Export
    Set exporter = LLExporter.Create(this.sourceWkb)
    filePath = exporter.ExportCustom(ExportNumber, folderPath, useFilt)

    appState.Restore

    ReportSavedFile filePath, exporter.LastExportPassword
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub


'@sub-title Get the sheet type tag from worksheet-level HiddenNames.
'@param sh Worksheet. The worksheet to query.
'@return String. The sheet type tag (HList, VList, etc.).
Private Function SheetTagOf(ByVal sh As Worksheet, _
                            Optional ByVal nameKey As String = "sheet_type") As String
    Dim shHn As HiddenNames
    Set shHn = HiddenNames.Create(sh)
    SheetTagOf = shHn.ValueAsString(nameKey)
End Function

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

'@label:update-filtered-data
'@sub-title Sync filtered data sheets before a filtered export
'@details
'Copies HList data to companion filtered sheets and removes hidden or
'empty rows, ensuring the filtered export reflects the user current
'filter state.
Private Sub UpdateFilteredData()

    Dim sh As Worksheet
    Dim filtsh As Worksheet
    Dim lo As ListObject
    Dim loRng As Range
    Dim destRng As Range
    Dim delRng As Range
    Dim filtLoHRng As Range
    Dim rowCounter As Long
    Dim wsFunc As WorksheetFunction

    On Error GoTo ErrUpdate

    Set wsFunc = Application.WorksheetFunction

    For Each sh In this.sourceWkb.Worksheets
        If SheetTagOf(sh) = "HList" Then
            Set lo = sh.ListObjects(1)

            If Not lo.DataBodyRange Is Nothing Then
                Set loRng = lo.DataBodyRange
                Set filtsh = this.sourceWkb.Worksheets(SheetTagOf(sh, "filtered_sheet"))
                rowCounter = loRng.Rows.Count

                With filtsh
                    On Error Resume Next
                    .ListObjects(1).DataBodyRange.Delete
                    On Error GoTo ErrUpdate

                    .ListObjects(1).Resize .Range(lo.Range.Address)
                    Set destRng = .Range(loRng.Address)
                    Set filtLoHRng = .ListObjects(1).HeaderRowRange
                    Set delRng = Nothing
                End With

                ' Copy values to filtered sheet
                destRng.Value = loRng.Value

                ' Mark hidden or empty rows for deletion (bottom-up)
                Do While rowCounter >= 1
                    If loRng.Cells(rowCounter, 1).EntireRow.Hidden Or _
                       wsFunc.CountA(filtLoHRng.Offset(rowCounter)) = 0 Then
                        If delRng Is Nothing Then
                            Set delRng = filtLoHRng.Offset(rowCounter)
                        Else
                            Set delRng = Application.Union(delRng, filtLoHRng.Offset(rowCounter))
                        End If
                    End If
                    rowCounter = rowCounter - 1
                Loop

                If Not delRng Is Nothing Then delRng.Delete
            End If
        End If
    Next sh
    Exit Sub

ErrUpdate:
    'MSG_ErrUpdate is in neither translation workbook, and TranslateTag answers
    'the tag itself on a miss, so this box used to show the raw key.
    MsgBox TranslatedValue("MSG_ErrExportData") & ": " & Err.Description, _
           vbCritical + vbOKOnly, TranslatedValue("MSG_Error")
End Sub

'@label:throw-error
'@sub-title Raise a ProjectError-based exception
'@param errNumber Long. A ProjectError enum value (e.g. ObjectNotInitialized).
'@param message String. Human-readable description of the failure.
'@throws ProjectError.<varies> Always raises the specified error.
Private Sub ThrowError(ByVal errNumber As Long, ByVal message As String)
    Err.Raise CLng(errNumber), CLASS_NAME, message
End Sub
