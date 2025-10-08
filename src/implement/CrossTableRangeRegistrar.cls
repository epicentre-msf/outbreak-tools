VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CrossTableRangeRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Centralises named range management for cross tables"

'@Folder("Implement")
'@ModuleDescription("Manage creation and refresh of cross table named ranges")

'@section State
'===============================================================================

Option Explicit

Private Type TCrossTableRangeRegistrar
    targetSheet As Worksheet
End Type

Private this As TCrossTableRangeRegistrar
Private Const CLASSNAME As String = "CrossTableRangeRegistrar"

Public Function Create(ByVal targetSheet As Worksheet) As CrossTableRangeRegistrar
    Dim registrar As CrossTableRangeRegistrar

    If targetSheet Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Target worksheet is required"
    End If

    Set registrar = New CrossTableRangeRegistrar
    Set registrar.TargetSheet = targetSheet
    Set Create = registrar
End Function

Public Property Set TargetSheet(ByVal sh As Worksheet)
    Set this.targetSheet = sh
End Property

Public Property Get TargetSheet() As Worksheet
    Set TargetSheet = this.targetSheet
End Property

'@section Public API
'===============================================================================

'@details Registers or refreshes a named range for the supplied layout summary.
Public Sub RegisterLayoutName(ByVal nameText As String, ByVal summary As CrossTableLayoutSummary)
    Dim targetRange As Range
    Dim workbook As Workbook
    Dim existingName As Name
    Dim refersTo As String

    If summary Is Nothing Then
        Err.Raise ProjectError.ObjectNotInitialized, CLASSNAME, "Layout summary is required"
    End If

    If LenB(Trim$(nameText)) = 0 Then
        Err.Raise ProjectError.InvalidArgument, CLASSNAME, "Name text cannot be empty"
    End If

    Set targetRange = ResolveTargetRange(nameText, summary)
    refersTo = "=" & targetRange.Address(True, True, xlA1, True)
    Set workbook = TargetSheet.Parent

    On Error Resume Next
        Set existingName = workbook.Names(nameText)
    On Error GoTo 0

    'Workbook names are preferred so downstream VBA can resolve them without sheet scope issues.
    If existingName Is Nothing Then
        workbook.Names.Add Name:=nameText, RefersTo:=refersTo
    Else
        existingName.RefersTo = refersTo
    End If
End Sub

'@details Removes a workbook/worksheet name if it exists.
Public Sub RemoveName(ByVal nameText As String)
    Dim workbook As Workbook

    If LenB(Trim$(nameText)) = 0 Then Exit Sub

    Set workbook = TargetSheet.Parent

    On Error Resume Next
        workbook.Names(nameText).Delete
        TargetSheet.Names(nameText).Delete
    On Error GoTo 0
End Sub

'@section Internal helpers
'===============================================================================

Private Function ResolveTargetRange(ByVal nameText As String, ByVal summary As CrossTableLayoutSummary) As Range
    Dim sheet As Worksheet
    Dim target As Range
    Dim startRow As Long
    Dim startColumn As Long
    Dim totalRows As Long
    Dim totalColumns As Long

    Set sheet = TargetSheet
    startRow = summary.StartRow
    startColumn = summary.StartColumn

    If Left$(nameText, 9) = "STARTROW_" Then
        Set target = sheet.Cells(startRow, startColumn)
    ElseIf Left$(nameText, 9) = "STARTCOL_" Then
        Dim headerRow As Long
        headerRow = startRow - 1
        If headerRow < 1 Then headerRow = 1
        Set target = sheet.Cells(headerRow, startColumn + 1)
    Else
        totalRows = summary.HeaderRowCount + summary.DataRowCount
        totalColumns = summary.DataColumnCount

        If totalRows <= 0 Then totalRows = 1
        If totalColumns <= 0 Then totalColumns = 1

        Set target = sheet.Cells(startRow, startColumn).Resize(totalRows, totalColumns)
    End If

    Set ResolveTargetRange = target
End Function

End Class
