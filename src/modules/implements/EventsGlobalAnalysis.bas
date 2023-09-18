Attribute VB_Name = "EventsGlobalAnalysis"
Attribute VB_Description = "Events associated to buttons and updates in all (uni, bi and time series analysis)"

'@Folder("Events")
'@ModuleDescription("Events associated to buttons and updates in all (uni, bi and time series analysis)")

Option Explicit
Option Private Module

Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"

Private tradsmess As ITranslation   'Translation of messages
Private lltrads As ILLTranslations
Private wb As Workbook


'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.cursor = xlDefault
End Sub

'Initialize translation of forms object
Private Sub InitializeTrads()
    Dim lltranssh As Worksheet
    Dim dicttranssh As Worksheet

    Set wb = ThisWorkbook
    Set lltranssh = wb.Worksheets(LLSHEET)
    Set dicttranssh = wb.Worksheets(TRADSHEET)

    'Those are private elms defined on to the top
    Set lltrads = LLTranslations.Create(lltranssh, dicttranssh)
    Set tradsmess = lltrads.TransObject()
End Sub

'@Description("Update the table which contains filters data in the linelist")
'@EntryPoint
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)
    Attribute UpdateFilterTables.VB_Description = "Update the table which contains filters data in the linelist"

    Dim sh As Worksheet                        'The actual worksheet
    Dim filtsh As Worksheet                    'Filtered worksheet
    Dim Lo As ListObject
    Dim filtCsTab As ICustomTable                'Filtered listObject custom table
    Dim destRng As Range
    Dim delRng As Range
    Dim LoRng As Range
    Dim rowCounter As Long
    Dim filtLoHrng As Range 'HeaderRowRange of listObject on filtered sheet

    On Error GoTo ErrUpdate

    InitializeTrads

    BusyApp cursor:=xlNorthwestArrow

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            'Clean the filtered table list object
            Set Lo = sh.ListObjects(1)

            If Not Lo.DataBodyRange Is Nothing Then
                Set LoRng = Lo.DataBodyRange
                Set filtsh = ThisWorkbook.Worksheets(sh.Cells(1, 5).Value)

                rowCounter = LoRng.Rows.Count

                With filtsh
                    On Error Resume Next
                        .ListObjects(1).DataBodyRange.Delete
                    On Error GoTo ErrUpdate

                    'Lo is the listObject
                    'LoRng is the listobject databodyrange in HList
                    'destRng is the listObject databodyrange in filtered sheet
                
                    .ListObjects(1).Resize .Range(Lo.Range.Address)
                    'This is the dataBodyRange of the filtered sheet
                    Set destRng = .Range(LoRng.Address)
                    Set filtLoHrng = .ListObjects(1).HeaderRowRange
                    'Initialize the range to delete at the end of the table
                    Set delRng = Nothing
                    Set filtCsTab = CustomTable.Create(.ListObjects(1))
                End With

                'move values to filtered sheet
                destRng.Value = LoRng.Value

                Do While rowCounter >= 1
                    If LoRng.Cells(rowCounter, 1).EntireRow.HIDDEN Then
                        If delRng Is Nothing Then
                            Set delRng = filtLoHrng.Offset(rowCounter)
                        Else
                            Set delRng = Application.Union(delRng, filtLoHrng.Offset(rowCounter))
                        End If
                    End If
                    rowCounter = rowCounter - 1
                Loop
                'Delete the range if necessary
                 If Not (delRng Is Nothing) Then delRng.Delete

                'Resize the custom table on the filtered worksheet (remove completly empty rows)
                filtCsTab.RemoveRows
            End If
        End If
    Next

    'caclulate active sheet
    DoEvents

    If calculate Then
        ActiveSheet.calculate
        ActiveSheet.UsedRange.calculate
        ActiveSheet.Columns("A:E").calculate
    End If

    NotBusyApp
    Exit Sub

ErrUpdate:
    MsgBox tradsmess.TranslatedValue("MSG_ErrUpdate") & ": " & Err.Description, vbCritical + vbOKOnly
    NotBusyApp
End Sub

'@Description("Find the selected column on "GOTO" Area and go to that column")
'@EntryPoint
Public Sub EventValueChangeAnalysis(Target As Range)

    Dim rng As Range
    Dim rngLook As Range
    Dim sLabel As String
    Dim actsh As Worksheet
    Dim analysisType As String
    Dim goToSection As String
    Dim goToHeader As String
    Dim goToGraph As String
    Dim rngName As String

    'Initialize translations (used in lltrads)
    InitializeTrads

    'Range name if it exists
    On Error Resume Next
        rngName = Target.Name.Name
    On Error GoTo 0

    On Error GoTo Err
    Set actsh = ActiveSheet

    analysisType = actsh.Cells(1, 3).Value

    Select Case analysisType

    Case "Uni-Bi-Analysis"
        'GoTo section range for univariate and bivariate analysis
        Set rng = actsh.Range("ua_go_to_section")

    Case "TS-Analysis", "SPT-Analysis"
        actsh.calculate
        actsh.UsedRange.calculate
        actsh.Columns("A:E").calculate
        'Goto section range for time series analysis
        If InStr(1, rngName, "ts_go_to_section") > 0 Then Set rng = Target
        'GoTo section range for spatio temporal analysis
        If InStr(1, rngName, "spt_go_to_section") > 0 Then Set rng = Target

    Case "SP-Analysis"
        'GoTo section for spatial analysis

        'The following events are in EventsSpatialAnalysis.bas.
        'They are triggered on tables or type geo.
        Set rng = actsh.Range("sp_go_to_section")
        If InStr(1, rngName, "ADM_DROPDOWN_") > 0 Then UpdateSingleSpTable rngName
        If InStr(1, rngName, "POPFACT_") > 0 Then DevideByPopulation rngName
        If InStr(1, rngName, "DEVIDEPOP_") > 0 Then FormatDevidePop rngName

    End Select

    If (Not (Intersect(Target, rng) Is Nothing)) And (Not rng Is Nothing) Then
        
        goToSection = lltrads.Value("gotosection")
        goToHeader = lltrads.Value("gotoheader")
        goToGraph = lltrads.Value("gotograph")

        sLabel = Replace(Target.Value, goToSection & ": ", vbNullString)
        sLabel = Replace(sLabel, goToHeader & ": ", vbNullString)
        sLabel = Replace(sLabel, goToGraph & ": ", vbNullString)

        Debug.Print sLabel
        Set rngLook = ActiveSheet.Cells.Find(What:=sLabel, LookIn:=xlValues, lookAt:=xlWhole, _
                                             MatchCase:=True, SearchFormat:=False)

        If Not rngLook Is Nothing Then rngLook.Activate
    End If

    Exit Sub
Err:
End Sub


'@Description("Intercept double click on a Spatio-temporal analysis sheet")
'@EntryPoint
Public Sub EventDoubleClickAnalysis(Target As Range)
    
    Dim rngName As String
    Dim sheetTag As String
    Dim actsh As Worksheet

    On Error Resume Next
        rngName = Target.Name.Name
    On Error GoTo Err

    Set actsh = ActiveSheet
    sheetTag = actsh.Cells(1, 3).Value
    If sheetTag <> "SPT-Analysis" Then Exit Sub
    
    If (InStr(1, rngName, "INPUTSPTGEO_") > 0) Then
        LoadGeo 0
    ElseIf (InStr(1, rngName, "INPUTSPTHF_") > 0) Then
        LoadGeo 1
    End If
Err:
End Sub