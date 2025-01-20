Attribute VB_Name = "EventsGlobal"
Option Explicit
'@Folder("Events")

'Module for all the global events: At the Workbook level
Private Const UPDATEDSHEETNAME As String = "__updated"
Private Const ANALYSISSHEET As String = "Analysis"
Private Const TRANSLATIONSHEET As String = "Translations"

'speed app
Private Sub BusyApp()
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
End Sub


'Check update status when something changes in a range on a worksheet
Public Sub checkUpdateStatus(ByVal sh As Worksheet, ByVal Target As Range)
    Dim upsh As Worksheet
    Dim upObj As IUpdatedValues
    Dim upId As String
    Dim Lo As ListObject
    Dim wb As Workbook

    BusyApp

    Set wb = ThisWorkbook

    Set upsh = wb.Worksheets(UPDATEDSHEETNAME)
    upId = LCase(Left(sh.Name, 4))
    If sh.Name = "Analysis" Then
        For Each Lo In sh.ListObjects
            upId = LCase(Replace(Lo.Name, "Tab_", vbNullString))
            Set upObj = UpdatedValues.Create(upsh, upId)
            upObj.CheckUpdate sh, Target
        Next
    Else
        Set upObj = UpdatedValues.Create(upsh, upId)
        upObj.CheckUpdate sh, Target
    End If

End Sub

'Fire this event when the workbook is opened
'@EntryPoint
Public Sub OpenedWorkbook()
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim rng As Range

    BusyApp
    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(ANALYSISSHEET)
    sh.Calculate
    On Error Resume Next
    Set rng = wb.Worksheets(TRANSLATIONSHEET).Range("RNG_NbTimesTrans")
    rng.Value = 0
    On Error GoTo 0
    'Set all update ranges to no
    SetAllUpdatedTo "yes"
    Application.CalculateBeforeSave = False
End Sub

Public Sub SetAllUpdatedTo(ByVal Value As String)
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim Lo As ListObject
    Dim rng As Range
    Dim counter As Long

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets(UPDATEDSHEETNAME)

    For Each Lo In sh.ListObjects
        Set rng = Lo.ListColumns("updated").Range
        For counter = 2 To rng.Rows.Count
            rng.Cells(counter, 1).Value = Value
        Next
    Next
End Sub

'Filter the checking worksheet on some of the checkings (Error, Warning, Notes, etcs)
Public Sub FilterCheckingsSheet(ByVal Target As Range)
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim filterValue As String
    Dim sheetRng As Range 'used range in the worksheet
    Dim cellRng As Range 'corresponding cellRange

    Set wb = ThisWorkbook
    Set sh = wb.Worksheets("__checkRep")
    Set sheetRng = sh.UsedRange
    If Intersect(Target, sh.Range("RNG_CheckingFilter")) Is Nothing Then Exit Sub
    filterValue = Target.Value

    If filterValue = vbNullString Then Exit Sub
    sh.Cells.EntireRow.Hidden = False

    If filterValue <> "All" Then
        Set cellRng = sheetRng.Cells(sheetRng.Rows.Count, 1)
        Do While cellRng.Row > sheetRng.Row
            'Hide cells with values corresponding to those selected (keeping headers)
            If (cellRng.Value <> filterValue) And (Not cellRng.Font.Size = 14) And (Not cellRng.Font.Size = 12) _
             And (Not cellRng.Value = vbNullString) Then cellRng.EntireRow.Hidden = True
            Set cellRng = cellRng.Offset(-1)
        Loop
    End If
End Sub

