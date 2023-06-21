Attribute VB_Name = "EventsGlobalAnalysis"
Attribute VB_Description = "Events associated to buttons and updates in all (uni, bi and time series analysis)"
Option Explicit
Option Private Module

Private Const LLSHEET As String = "LinelistTranslation"
Private Const TRADSHEET As String = "Translations"
Private Const PASSSHEET As String = "__pass"

Private tradsmess As ITranslation   'Translation of messages
Private pass As ILLPasswords
Private lltrads As ILLTranslations
Private wb As Workbook

'Subs to speed up the application
'speed app
Private Sub BusyApp(Optional ByVal cursor As Long = xlDefault)
    Application.ScreenUpdating = False
    Application.EnableAnimations = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = cursor
End Sub

'Return back to previous state
Private Sub NotBusyApp()
    Application.ScreenUpdating = True
    Application.EnableAnimations = True
    Application.Cursor = xlDefault
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
    Set pass = LLPasswords.Create(wb.Worksheets(PASSSHEET))
End Sub

'@Description("Update the table which contains filters data in the linelist")
'@EntryPoint
Public Sub UpdateFilterTables(Optional ByVal calculate As Boolean = True)

    Dim sh As Worksheet                        'The actual worksheet
    Dim filtsh As Worksheet                    'Filtered worksheet
    Dim Lo As ListObject
    Dim destRng As Range
    Dim delRng As Range
    Dim LoRng As Range
    Dim rowCounter As Long
    Dim filtLoHrng As Range 'HeaderRowRange of listObject on filtered sheet

    On Error GoTo ErrUpdate

    InitializeTrads

    BusyApp cursor:= xlNorthwestArrow

    For Each sh In wb.Worksheets
        If sh.Cells(1, 3).Value = "HList" Then
            'Clean the filtered table list object
            Set Lo = sh.ListObjects(1)

            If Not Lo.DataBodyRange Is Nothing Then
                Set LoRng = Lo.DataBodyRange
                Set filtsh = ThisWorkbook.Worksheets(sh.Cells(1, 5).Value)

                rowCounter = LoRng.Rows.Count

                On Error Resume Next
                    filtsh.ListObjects(1).DataBodyRange.Delete
                On Error GoTo ErrUpdate

                'Lo is the listObject
                'LoRng is the listobject databodyrange in HList
                'destRng is the listObject databodyrange in filtered sheet
                With filtsh
                    .ListObjects(1).Resize .Range(Lo.Range.Address)
                    'This is the dataBodyRange of the filtered sheet
                    Set destRng = .Range(LoRng.Address)
                    Set filtLoHrng = .ListObjects(1).HeaderRowRange
                    'Initialize the range to delete at the end of the table
                    Set delRng = Nothing
                End With

                'move values to filtered sheet
                destRng.Value = LoRng.Value

                Do While rowCounter >= 1
                    If LoRng.Cells(rowCounter, 1).EntireRow.Hidden Then 
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
Sub EventValueChangeAnalysis(Target As Range)

    Dim rng As Range
    Dim RngLook As Range
    Dim sLabel As String
    Dim actSh As Worksheet
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
    Set actSh = ActiveSheet

    analysisType = actSh.Cells(1, 3).Value

    Select Case analysisType

    Case "Uni-Bi-Analysis"
        'GoTo section range for univariate and bivariate analysis
        Set rng = actSh.Range("ua_go_to_section")

    Case "TS-Analysis"
        actSh.calculate
        actSh.UsedRange.calculate
        actSh.Columns("A:E").calculate
        'Goto section range for time series analysis
        If InStr(1, rngName, "ts_go_to_section") > 0 Then Set rng = Target

    Case "SP-Analysis"
        'GoTo section for spatial analysis

        'The following events are in EventsSpatialAnalysis.bas.
        'They are triggered on tables or type geo.
        Set rng = actSh.Range("sp_go_to_section")
        If InStr(1, rngName, "ADM_DROPDOWN_") > 0 Then UpdateSingleSpTable rngName
        If InStr(1, rngName, "POPFACT_") > 0 Then DevideByPopulation rngName
        If InStr(1, rngName, "DEVIDEPOP_") > 0 Then FormatDevidePop rngName

    End Select

    If (Not (Intersect(Target, rng) Is Nothing)) And (Not rng Is Nothing) Then
        
        goToSection = lltrads.Value("gotosection")
        goToHeader = lltrads.Value("gotoheader")
        goToGraph = lltrads.Value("gotograph")

        sLabel = Replace(Target.Value, goToSection & ": ", "")
        sLabel = Replace(sLabel, goToHeader & ": ", "")
        sLabel = Replace(sLabel, goToGraph & ": ", "")

        Debug.Print sLabel
        Set RngLook = ActiveSheet.Cells.Find(What:=sLabel, LookIn:=xlValues, LookAt:=xlWhole, _
                                             MatchCase:=True, SearchFormat:=False)

        If Not RngLook Is Nothing Then RngLook.Activate
    End If

    Exit Sub
Err:
End Sub