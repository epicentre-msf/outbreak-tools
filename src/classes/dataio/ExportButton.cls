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
'@depends ILLExporter, LLExporter, IApplicationState, ApplicationState, IOSFiles, OSFiles

Option Explicit


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
    trads As ITranslationObject
End Type

Private this As TExportButton


'@section Factory
'===============================================================================

' @description Create an export button wrapper.
' @param sourceWkb Workbook. The linelist workbook.
' @param trads ITranslationObject. Translations for messages.
' @param expBtn MSForms.CommandButton. The button to listen to.
' @param expChk MSForms.CheckBox. Optional filter checkbox.
' @return ExportButton wired to the button.
' @export
Public Function Create(ByVal sourceWkb As Workbook, _
                       ByVal trads As ITranslationObject, _
                       ByVal expBtn As MSForms.CommandButton, _
                       Optional ByVal expChk As MSForms.CheckBox = Nothing) As ExportButton

    If sourceWkb Is Nothing Then ThrowError "Workbook required"
    If trads Is Nothing Then ThrowError "Translations required"
    If expBtn Is Nothing Then ThrowError "Button required"

    With New ExportButton
        Set .SourceWorkbook = sourceWkb
        Set .Translations = trads
        Set .PushButton = expBtn
        If Not expChk Is Nothing Then Set .FilterCheckBox = expChk
        Set Create = .Self
    End With
End Function


'@section Public accessors
'===============================================================================

Public Property Get Self() As ExportButton
    Set Self = Me
End Property

Public Property Set SourceWorkbook(ByVal wb As Workbook)
    Set this.sourceWkb = wb
End Property

Public Property Set Translations(ByVal trads As ITranslationObject)
    Set this.trads = trads
End Property

Public Property Set PushButton(ByVal expBtn As MSForms.CommandButton)
    Set Btn = expBtn
End Property

Public Property Set FilterCheckBox(ByVal expChk As MSForms.CheckBox)
    Set chkFilt = expChk
End Property

' @description Export number derived from button name (e.g. CMDExport1 -> 1).
Public Property Get ExportNumber() As Long
    ExportNumber = CLng(Replace(Btn.Name, NAMETAG, vbNullString))
End Property

' @description Whether the filter checkbox is checked.
Public Property Get UseFilter() As Boolean
    If Not chkFilt Is Nothing Then UseFilter = chkFilt.Value
End Property

Public Property Let UseFilter(ByVal value As Boolean)
    If Not chkFilt Is Nothing Then chkFilt.Value = value
End Property


'@section Event handler
'===============================================================================

Private Sub Btn_Click()
    RunExport
End Sub

' @description Perform a custom export.
' Selects a folder, prepares filters if needed, delegates to LLExporter.
Private Sub RunExport()

    Dim exporter As ILLExporter
    Dim appState As IApplicationState
    Dim io As IOSFiles
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
    appState.ApplyBusyState True, False, xlWait, False

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
    Exit Sub

ErrHand:
    On Error Resume Next
    MsgBox TranslatedValue("MSG_ErrHandExport"), _
           vbOKOnly + vbCritical, TranslatedValue("MSG_Error")
    If Not exporter Is Nothing Then exporter.CloseAll
    If Not appState Is Nothing Then appState.Restore
End Sub


'@section Private helpers
'===============================================================================

' @description Get a translated message value.
Private Function TranslatedValue(ByVal msgCode As String) As String
    TranslatedValue = this.trads.TranslatedValue(msgCode)
End Function

' @description Sync filtered data sheets before a filtered export.
' Copies HList data to companion filtered sheets, removing hidden rows.
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
        If CStr(sh.Cells(1, 3).Value) = "HList" Then
            Set lo = sh.ListObjects(1)

            If Not lo.DataBodyRange Is Nothing Then
                Set loRng = lo.DataBodyRange
                Set filtsh = this.sourceWkb.Worksheets(CStr(sh.Cells(1, 5).Value))
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
    MsgBox TranslatedValue("MSG_ErrUpdate") & ": " & Err.Description, _
           vbCritical + vbOKOnly
End Sub

' @description Raise a class-level exception.
Private Sub ThrowError(ByVal message As String)
    Err.Raise VBA.vbObjectError + 516, CLASS_NAME, message
End Sub
