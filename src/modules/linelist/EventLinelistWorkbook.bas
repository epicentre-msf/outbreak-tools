Attribute VB_Name = "EventLinelistWorkbook"
Option Explicit

'@Folder("Linelist Events")
'@ModuleDescription("Thin workbook-level event handlers delegating to the shared LinelistEventsManager")
'@IgnoreModule UnrecognizedAnnotation, SuperfluousAnnotationArgument, ExcelMemberMayReturnNothing, UseMeaningfulName, HungarianNotation

Private mBooting As Boolean

Private Sub Workbook_Open()
    Application.ScreenUpdating = False

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.LLWorkbookOpened

Clean:
    mBooting = False
End Sub

' A sheet deactivation is answered by Worksheet_Deactivate in the code module of
' each HList sheet, which is the only kind of sheet that has anything to do
' there. EventLinelistSheet carries that handler.

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    LinelistEventsManager.LLSheetChanged sh, target

Clean:
    mBooting = False
End Sub

' A double-click on a spatial analysis sheet opens one of the two geo pickers.
' The picker is a UserForm, GeoModule is what shows it, and the call belongs
' here. EventLinelist answers which of the two the click asked for.
Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As Range, Cancel As Boolean)
    Dim geoScopeWanted As Long

    If mBooting Then Exit Sub
    If TypeName(sh) <> "Worksheet" Then Exit Sub

    mBooting = True

    On Error GoTo Clean
    geoScopeWanted = LinelistEventsManager.DoubleClicked(sh, target)
    If geoScopeWanted >= 0 Then GeoModule.LoadGeo CByte(geoScopeWanted)

Clean:
    mBooting = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Application.FormatStaleValues = True
    LinelistEventsManager.DisposeEventLinelist
    On Error GoTo 0
End Sub

' Save As is held to the macro-enabled binary format. A linelist saved as .xlsx
' comes back with every module, form and button gone, and the user has no way to
' tell until a button does nothing.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim savePath As String

    Application.CalculateBeforeSave = False

    If Not SaveAsUI Then Exit Sub

    Cancel = True

    savePath = AskForSavePath()
    If LenB(savePath) = 0 Then Exit Sub

    On Error Resume Next
    Application.EnableEvents = False
    ThisWorkbook.SaveAs Filename:=savePath, FileFormat:=xlExcel12
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' Ask the user where to save, and answer a path that ends in .xlsb.
'
' Mac Excel refuses the two-part FileFilter string Windows takes -- "Excel
' Binary Workbook (*.xlsb), *.xlsb" -- and the whole call raises "Method
' GetSaveAsFilename of object _Application failed", so Save As did nothing at
' all on a Mac. The dialog is asked for three ways, from the richest to the
' plainest, and the first one that answers wins.
'
' The extension is put back on whatever comes out. The filter is what held the
' user to .xlsb, and a dialog opened without one hands back whatever was typed.
Private Function AskForSavePath() As String
    Dim answer As Variant
    Dim startName As String

    startName = ThisWorkbook.Name

    On Error Resume Next

    answer = Application.GetSaveAsFilename( _
        InitialFileName:=startName, _
        FileFilter:="Excel Binary Workbook (*.xlsb), *.xlsb", _
        Title:="Save As")

    If Err.Number <> 0 Then
        Err.Clear
        answer = Application.GetSaveAsFilename( _
            InitialFileName:=startName, _
            FileFilter:="*.xlsb", _
            Title:="Save As")
    End If

    If Err.Number <> 0 Then
        Err.Clear
        answer = Application.GetSaveAsFilename(InitialFileName:=startName)
    End If

    On Error GoTo 0

    ' A cancelled dialog answers the Boolean False, and a dialog that never
    ' opened leaves the variant empty.
    If VarType(answer) = vbBoolean Then Exit Function
    If IsEmpty(answer) Then Exit Function

    AskForSavePath = WithBinaryExtension(CStr(answer))
End Function

' Put the .xlsb extension on a path, whatever the user typed.
Private Function WithBinaryExtension(ByVal filePath As String) As String
    Const BINARY_EXTENSION As String = ".xlsb"

    Dim trimmedPath As String
    Dim dotAt As Long
    Dim separatorAt As Long

    trimmedPath = Trim$(filePath)
    If LenB(trimmedPath) = 0 Then Exit Function

    dotAt = InStrRev(trimmedPath, ".")
    separatorAt = InStrRev(trimmedPath, Application.PathSeparator)

    ' A dot inside a folder name is not an extension.
    If dotAt > separatorAt And dotAt > 0 Then
        trimmedPath = Left$(trimmedPath, dotAt - 1)
    End If

    WithBinaryExtension = trimmedPath & BINARY_EXTENSION
End Function
